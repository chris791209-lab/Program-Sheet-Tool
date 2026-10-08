import streamlit as st
import pandas as pd
import io
import os
import xlsxwriter
import tempfile
import re 
import openpyxl
from openpyxl_image_loader import SheetImageLoader
from openpyxl.utils import get_column_letter
from PIL import Image

# --- 網頁介面設定 ---
st.set_page_config(page_title="Program Sheet 生成器", layout="centered")
# ==========================================
# 欄位別名表 (單一來源)
# 卡片欄位 -> master 可能出現的欄名，由左到右找到有值的為止。
# 每季欄名會變 (例如結尾多一個 " *")，新欄名請加在這裡，不要寫死在下方。
# ==========================================
FIELD_ALIASES = {
    'dpci':         ['DPCI'],
    'style':        ['Manufacturer Style # *', 'Manufacturer Style #', 'Style Number'],
    'upc':          ['Barcode', 'UPC#', 'UPC'],
    'pid':          ['Spark PID', 'PID'],
    'desc':         ['Vendor Product Description *', 'Vendor Product Description', 'Product Description'],
    'fca':          ['FCA Factory City Unit Cost', 'FCA', 'FCA $'],
    'fob':          ['FOB Unit Cost', 'FOB', 'FOB $'],
    'retail':       ['Suggested Unit Retail', 'RETAIL', 'Retail$'],
    'packaging':    ['Retail Packaging Format (1) *', 'Retail Packaging Format (1)', 'Packaging'],
    'hs':           ['HTS Code', 'HS NO'],
    'case':         ['Case Unit Quantity', 'Casepack'],
    'inner':        ['Inner Pack Unit Quantity', 'Innerpack'],
    'material':     ['Primary Raw Material Type', 'Material', 'Main Raw Material *'],
    'qty':          ['Ent Ttl Rcpt U', 'Total Units', 'QTY'],
    'factory_name': ['Factory Name', 'Factory info.', 'Factory'],
    'factory_id':   ['Factory ID', 'Import Vendor ID'],
    'order_point':  ['Import Vendor Order Point'],
}
DEFAULT_PROGRAM_NAME = "202X D240  PROGRAM NAME"


def program_name_from_filename(filename):
    """
    由 PCN 檔名推出 Program 名稱，用於輸出檔名與表頭標題。
        '2027_C2_D240_Easter.xlsx'      -> '2027 C2 D240 Easter'
        '2027 C2 D240 Easter (1).xlsx'  -> '2027 C2 D240 Easter'
    """
    stem = os.path.splitext(os.path.basename(str(filename)))[0]
    stem = re.sub(r'\s*\(\d+\)\s*$', '', stem)      # 瀏覽器重複下載加上的 " (1)"
    stem = re.sub(r'[_\s]+', ' ', stem).strip()      # 底線 / 連續空白 -> 單一空白
    return stem or DEFAULT_PROGRAM_NAME


# 缺少時會跳出警告的欄位 (卡片標籤名稱)
REQUIRED_FIELDS = {
    'dpci': 'DPCI', 'style': 'Style', 'upc': 'UPC#', 'pid': 'PID', 'desc': 'Description',
    'retail': 'RETAIL', 'case': 'Casepack (Case)', 'inner': 'Casepack (Inner)', 'qty': 'QTY',
    'factory_name': 'Factory Name', 'factory_id': 'Factory ID', 'order_point': 'Order Point',
}
# 有些季別的 master 本來就沒有的欄位，缺少時只提示、不警告
OPTIONAL_FIELDS = {'packaging': 'Packaging', 'hs': 'HS NO', 'material': 'Material'}

# ==========================================
# 0. 密碼保護機制
# ==========================================
def check_password():
    """回傳 True 代表使用者輸入了正確的密碼"""
    def password_entered():
        if st.session_state["password"] == st.secrets["app_password"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]  
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.text_input(
            "🔒 請輸入 AE 部門共用密碼以啟用工具：", 
            type="password", 
            on_change=password_entered, 
            key="password"
        )
        return False
    
    elif not st.session_state["password_correct"]:
        st.text_input(
            "🔒 請輸入 AE 部門共用密碼以啟用工具：", 
            type="password", 
            on_change=password_entered, 
            key="password"
        )
        st.error("❌ 密碼錯誤，請重新輸入。")
        return False
    else:
        return True

if not check_password():
    st.stop()

# ========================================================
# 🚀 以下為登入後的主要系統介面
# ========================================================
st.title("🚀 Program Sheet 自動生成器")
st.markdown("請依序完成以下步驟，系統將瞬間為您篩選資料並排版整合！")

# ========================================================
# 📄 步驟 1：上傳商品資料 (PCN)
# ========================================================
st.markdown("### 📄 步驟 1：上傳商品資料 (PCN)")
uploaded_data = st.file_uploader("請上傳包含完整報價、工廠等資訊的 Excel 檔 (.xlsm / .xlsx)", type=["xlsm", "xlsx"])

if uploaded_data is not None:
    st.success("✅ 資料檔已讀取！")

    # Program 名稱：預設取自 PCN 檔名，同時用於下載檔名與每個頁籤左上角的標題
    program_name = st.text_input(
        "📝 Program 名稱 (自動取自 PCN 檔名，可修改；會用於下載檔名與表頭標題)",
        value=program_name_from_filename(getattr(uploaded_data, 'name', '')),
    )
    program_name = (program_name or '').strip() or DEFAULT_PROGRAM_NAME
    
    try:
        uploaded_data.seek(0)
        df = pd.read_excel(uploaded_data, sheet_name="Data", engine="openpyxl")
        df.columns = [str(c).strip() for c in df.columns]
    except Exception as e:
        uploaded_data.seek(0)
        df = pd.read_excel(uploaded_data, engine="openpyxl")
        df.columns = [str(c).strip() for c in df.columns]
        
    if 'DPCI' in df.columns:
        df = df.dropna(subset=['DPCI'])
    # --- Shipper (陳列道具) 排除：Program Sheet 只列商品，Shipper 屬於 PO GRID ---
    ship_mask = pd.Series(False, index=df.index)
    for c in FIELD_ALIASES['desc']:
        if c in df.columns:
            ship_mask = ship_mask | df[c].astype(str).str.strip().str.lower().str.startswith('shipper:')
    shipper_count = int(ship_mask.sum())
    if shipper_count:
        shipper_ids = "、".join(str(x) for x in df.loc[ship_mask, 'DPCI'].tolist()) if 'DPCI' in df.columns else ""
        exclude_shippers = st.checkbox(
            f"🚫 排除 {shipper_count} 筆 Shipper 陳列道具 (描述以「Shipper:」開頭，無圖片、無 QTY)",
            value=True,
            help=shipper_ids or None,
        )
        if exclude_shippers:
            df = df[~ship_mask].copy()

    # --- 欄位檢查：欄名每季都會變，找不到時要講出來，不能靜默留空 ---
    missing_fields = [label for key, label in REQUIRED_FIELDS.items()
                      if not any(c in df.columns for c in FIELD_ALIASES[key])]
    if not any(c in df.columns for c in FIELD_ALIASES['fca'] + FIELD_ALIASES['fob']):
        missing_fields.append('FCA $')
    if missing_fields:
        st.warning("⚠️ 資料檔中找不到以下卡片欄位對應的欄名，生成後這些欄位會是空白："
                   f"**{'、'.join(missing_fields)}**。若本季欄名有變動，請將新欄名加入程式上方的 FIELD_ALIASES。")
    absent_optional = [label for key, label in OPTIONAL_FIELDS.items()
                       if not any(c in df.columns for c in FIELD_ALIASES[key])]
    if absent_optional:
        st.caption(f"ℹ️ 本季資料檔沒有這些選填欄位，卡片上會留空：{'、'.join(absent_optional)}")

    # ========================================================
    # 🗂️ 步驟 2：資料動態篩選
    # ========================================================
    st.markdown("### 🗂️ 步驟 2：選擇產出類別 (部門分工用)")
    st.info("💡 系統已鎖定分類欄位。若只需產出特定類別供部門分工，請在此篩選；若留空則會產出全部商品。")
    
    target_cols = ['Subclass Name', 'Factory Name']
    available_cols = [c for c in target_cols if c in df.columns]
    
    if available_cols:
        filter_col = st.selectbox("📌 1. 請確認用來分類的「欄位名稱」：", available_cols)
    else:
        st.warning("⚠️ 檔案中找不到 'Subclass Name' 或 'Factory Name'，已顯示所有欄位供您選擇。")
        filter_col = st.selectbox("📌 1. 請確認用來分類的「欄位名稱」：", df.columns)

    unique_vals = sorted([str(x).strip() for x in df[filter_col].dropna().unique() if str(x).strip() != ''])
    selected_categories = st.multiselect(f"🎯 2. 請選擇要輸出的【{filter_col}】(可多選)：", unique_vals)

    if selected_categories:
        df_filtered = df[df[filter_col].astype(str).str.strip().isin(selected_categories)].copy()
        st.write(f"📊 目前已篩選：將產出 **{len(df_filtered)}** 筆商品 (原始總數 {len(df)} 筆)")
    else:
        df_filtered = df.copy()
        st.write(f"📊 目前未套用篩選：將產出 **全部 {len(df_filtered)}** 筆商品")

    # ========================================================
    # 🖼️ 步驟 3：上傳圖片來源
    # ========================================================
    st.markdown("### 🖼️ 步驟 3：上傳圖片來源")
    uploaded_img = st.file_uploader("📊 上傳有縮圖的 Excel 檔 (SPK)", type=["xlsx", "xlsm"])

    # ========================================================
    # ⚙️ 步驟 4：執行生成
    # ========================================================
    st.markdown("### ⚙️ 步驟 4：執行生成")
    if st.button("✨ 生成 Program Sheet"):
        
        if df_filtered.empty:
            st.error("⚠️ 篩選後的資料為空，無法生成！請重新調整您的篩選條件。")
            st.stop()
            
        with st.spinner("正在為您進行排版與處理圖片，請稍候 (若包含萃取圖片可能需要數十秒)..."):
            temp_dir = tempfile.mkdtemp()
            
            if uploaded_img is not None:
                try:
                    wb = openpyxl.load_workbook(uploaded_img, data_only=True)
                    sheet_name_to_use = "Products" if "Products" in wb.sheetnames else ("Data" if "Data" in wb.sheetnames else wb.sheetnames[0])
                    sheet = wb[sheet_name_to_use]
                    image_loader = SheetImageLoader(sheet)

                    header_row, thumb_col, dpci_col = None, None, None
                    for r in range(1, 15): 
                        for c in range(1, sheet.max_column + 1):
                            val = str(sheet.cell(row=r, column=c).value).strip().lower()
                            if val == 'thumbnail':
                                thumb_col = c
                                header_row = r
                            elif val in ['dpci', 'pid', 'spark pid']:
                                dpci_col = c
                        if thumb_col and dpci_col:
                            break

                    if thumb_col and dpci_col:
                        thumb_letter = get_column_letter(thumb_col)
                        for r in range(header_row + 1, sheet.max_row + 1):
                            dpci_val = str(sheet.cell(row=r, column=dpci_col).value).strip()
                            img_cell = f"{thumb_letter}{r}"

                            if dpci_val and dpci_val.lower() not in ['none', 'nan', '']:
                                if image_loader.image_in(img_cell):
                                    try:
                                        img = image_loader.get(img_cell)
                                        safe_name = "".join(x for x in dpci_val if x.isalnum() or x in "-_")
                                        img_path = os.path.join(temp_dir, f"{safe_name}.png")
                                        img.save(img_path, "PNG")
                                    except Exception:
                                        pass
                    else:
                        st.warning("⚠️ 在 SPK 檔中找不到 'Thumbnail' 或 DPCI / PID 標題，本次不會放入圖片。")
                except Exception as e:
                    st.warning(f"⚠️ 解析 SPK 圖片失敗: {e}")
            
            output = io.BytesIO()
            workbook = xlsxwriter.Workbook(output, {'in_memory': True})
            
            base_props = {
                'label': {'font_name': 'Arial', 'font_size': 10, 'bold': True, 'align': 'left', 'valign': 'vcenter'},
                'data': {'font_name': 'Arial', 'font_size': 10, 'align': 'left', 'valign': 'vcenter', 'text_wrap': True},
                'img': {'bg_color': '#FFFFFF'}
            }
            
            fmt = {}
            def create_fmt(name, base, **kwargs):
                p = base_props[base].copy()
                p.update(kwargs)
                fmt[name] = workbook.add_format(p)
                
            create_fmt('img_top', 'img', top=2, left=2, right=2, bottom=1)
            create_fmt('lbl_l', 'label', left=2)
            create_fmt('lbl_in', 'label')
            create_fmt('lbl_lb', 'label', left=2, bottom=2)
            create_fmt('dat_in', 'data')
            create_fmt('dat_r', 'data', right=2)
            create_fmt('dat_b', 'data', bottom=2)
            create_fmt('dat_rb', 'data', right=2, bottom=2)
            
            create_fmt('dat_in_curr', 'data', num_format='$#,##0.00')
            create_fmt('dat_r_curr', 'data', right=2, num_format='$#,##0.00')
            create_fmt('fact_merge', 'label', left=2, right=2, bottom=2, bg_color='#FFC000')

            create_fmt('hdr_title', 'label', font_size=13, align='left')
            # Program 名稱長度每季不同，過長時自動縮小字級，避免被右側 CATEGORY 欄截斷
            create_fmt('hdr_title_name', 'label', font_size=13, align='left', shrink=True)
            create_fmt('hdr_lbl', 'label', align='left')
            create_fmt('hdr_input', 'data', align='left', bg_color='#F4F4F4', border=1)

            # 生成過程中發現的資料狀況，結束後統一顯示 (只在 Master Sheet 那一輪記錄)
            issues = {'fob': [], 'order_point': [], 'qty_fraction': [], 'no_image': [], 'card_error': []}

            def get_val(row_series, possible_cols):
                for col in possible_cols:
                    if col in row_series.index:
                        raw = row_series[col]
                        # 整數被 Excel/pandas 讀成浮點數時 (例如 397959.0)，還原成整數再轉字串
                        if isinstance(raw, float) and not pd.isna(raw) and raw.is_integer():
                            raw = int(raw)
                        val = str(raw).strip()
                        if val.lower() not in ['nan', '', 'nat', 'none']:
                            return val
                return ""

            def to_float(val):
                if not val: return ""
                try: return float(val.replace('$', '').replace(',', '').strip())
                except ValueError: return val

            def to_int_str_comma(val):
                # 必須四捨五入：Excel 會把 125454 存成 125453.99999999999，
                # 直接 int() 會無條件捨去而少 1。
                if not val: return ""
                try: return f"{int(round(float(str(val).replace(',', '').strip()))):,}"
                except ValueError: return str(val)

            def order_point_suffix(val):
                """
                取 Import Vendor Order Point 最後一個連字號後面的數字。
                    'Natty Boy - 240-20'                                  -> '20'
                    '397959 - Ganzhou Keshou Arts & Crafts Co., Ltd.-62'  -> '62'
                    'Wenzhou -240-8'                                      -> '8'
                舊寫法取最後兩個字元，個位數會變成 '-8'，三位數會少一碼。
                """
                val = str(val).strip()
                m = re.search(r'-\s*(\d+)\s*$', val)
                if m:
                    return m.group(1)
                if re.fullmatch(r'\d+', val):
                    return val
                return ""
                
            def w_row(ws, s_row, s_col, r_off, c0, c1, c2, c3, c4, f0, f1, f2, f3, f4):
                ws.write(s_row + r_off, s_col + 0, c0, fmt[f0])
                ws.write(s_row + r_off, s_col + 1, c1, fmt[f1])
                ws.write(s_row + r_off, s_col + 2, c2, fmt[f2]) 
                ws.write(s_row + r_off, s_col + 3, c3, fmt[f3])
                ws.write(s_row + r_off, s_col + 4, c4, fmt[f4])

            df_filtered['RawFactoryName'] = df_filtered.apply(lambda r: get_val(r, FIELD_ALIASES['factory_name']), axis=1)
            df_filtered = df_filtered.sort_values(by='RawFactoryName').reset_index(drop=True)

            def draw_cards_on_sheet(ws, current_df, track=False):
                ws.set_landscape()
                ws.set_margins(left=0.3, right=0.3, top=0.4, bottom=0.4)
                ws.fit_to_pages(1, 0)
                
                for i in range(3):
                    base = i * 6
                    ws.set_column(base, base, 13)       
                    ws.set_column(base + 1, base + 1, 22) 
                    ws.set_column(base + 2, base + 2, 2)  
                    ws.set_column(base + 3, base + 3, 13) 
                    ws.set_column(base + 4, base + 4, 22) 
                    ws.set_column(base + 5, base + 5, 4)  

                ws.set_row(0, 30)
                ws.merge_range(0, 0, 0, 1, f"{program_name} - ", fmt['hdr_title_name'])
                ws.merge_range(0, 2, 0, 4, "CATEGORY", fmt['hdr_title'])
                ws.data_validation(0, 2, 0, 4, {'validate': 'list', 'source': ['Decor', 'Kids Activity', 'Party', 'Giveaways', 'TOT']})

                ws.set_row(1, 20)
                ws.merge_range(1, 0, 1, 1, "Business award date:", fmt['hdr_lbl'])
                ws.merge_range(1, 2, 1, 4, "", fmt['hdr_input']) 
                ws.write(1, 6, "Vendor ID#:", fmt['hdr_lbl'])                 
                ws.merge_range(1, 7, 1, 9, "", fmt['hdr_input'])

                ws.set_row(2, 20)
                ws.merge_range(2, 0, 2, 1, "Sourcing:", fmt['hdr_lbl'])
                ws.merge_range(2, 2, 2, 4, "", fmt['hdr_input'])
                ws.data_validation(2, 2, 2, 4, {'validate': 'list', 'source': ['Christy Meyers - Van Der Bosch', 'Ashley Krucker', 'Angela Kennedy']})
                ws.write(2, 6, "PD&D:", fmt['hdr_lbl'])                       
                ws.merge_range(2, 7, 2, 9, "", fmt['hdr_input'])              
                ws.data_validation(2, 7, 2, 9, {'validate': 'list', 'source': ['Adam Hoppus', 'La Dieh Rosenthal', 'Name D']})

                ws.set_row(3, 20)
                ws.merge_range(3, 0, 3, 1, "TSS MR:", fmt['hdr_lbl'])
                ws.merge_range(3, 2, 3, 4, "", fmt['hdr_input'])
                ws.data_validation(3, 2, 3, 4, {'validate': 'list', 'source': ['Asya Yi', 'Jasmine Li', 'Feng Cao']})
                ws.write(3, 6, "Set date:", fmt['hdr_lbl'])                   
                ws.merge_range(3, 7, 3, 9, "", fmt['hdr_input'])

                item_index = 0
                page_breaks = [] 
                
                for index, row in current_df.iterrows():
                    try:
                        block_row = item_index // 3
                        block_col = item_index % 3
                        
                        start_row = 5 + (block_row * 13)
                        start_col = block_col * 6
                        
                        ws.set_row(start_row, 234) 
                        ws.merge_range(start_row, start_col, start_row, start_col + 4, "", fmt['img_top'])
                        
                        for r in range(start_row + 1, start_row + 11):
                            ws.set_row(r, 22)
                        
                        if block_row > 0 and block_row % 2 == 0 and block_col == 0:
                            page_breaks.append(start_row - 1)
                        
                        dpci_val = get_val(row, FIELD_ALIASES['dpci'])
                        style_val = get_val(row, FIELD_ALIASES['style'])
                        upc_val = get_val(row, FIELD_ALIASES['upc'])
                        pid_val = get_val(row, FIELD_ALIASES['pid'])
                        desc_val = get_val(row, FIELD_ALIASES['desc'])

                        try:
                            if upc_val:
                                upc_val = str(int(float(upc_val)))
                        except Exception:
                            pass

                        # 成本：優先 FCA；FCA 空白 (FOB-only 商品) 時改用 FOB，並列入提醒
                        cost_raw = get_val(row, FIELD_ALIASES['fca'])
                        cost_label = "FCA $:"
                        if not cost_raw:
                            cost_raw = get_val(row, FIELD_ALIASES['fob'])
                            if cost_raw:
                                cost_label = "FOB $:"   # 標籤跟著成本基礎走，避免 FOB 金額掛在 FCA 標籤下
                                if track:
                                    issues['fob'].append(dpci_val)
                        fca_val = to_float(cost_raw)
                        retail_val = to_float(get_val(row, FIELD_ALIASES['retail']))
                        pack_val = get_val(row, FIELD_ALIASES['packaging'])
                        hs_val = get_val(row, FIELD_ALIASES['hs'])

                        case_q = get_val(row, FIELD_ALIASES['case'])
                        inner_q = get_val(row, FIELD_ALIASES['inner'])
                        try: case_q = str(int(float(case_q))) if case_q else ""
                        except: pass
                        try: inner_q = str(int(float(inner_q))) if inner_q else ""
                        except: pass
                        pack_str = f"{case_q} / {inner_q}" if (case_q or inner_q) else ""

                        mat_val = get_val(row, FIELD_ALIASES['material'])
                        qty_raw = get_val(row, FIELD_ALIASES['qty'])
                        qty_val = to_int_str_comma(qty_raw)
                        # QTY 理論上是整數。極小的浮點誤差 (…99999999999) 直接四捨五入吸收；
                        # 若真的帶小數，代表 master 資料有問題，要提醒而不是默默進位。
                        try:
                            q = float(qty_raw.replace(',', ''))
                            if abs(q - round(q)) > 0.001 and track:
                                issues['qty_fraction'].append(f"{dpci_val}（{qty_raw}）")
                        except ValueError:
                            pass

                        raw_factory_name = get_val(row, FIELD_ALIASES['factory_name'])
                        raw_factory_id = get_val(row, FIELD_ALIASES['factory_id'])
                        raw_vendor_op = get_val(row, FIELD_ALIASES['order_point'])
                        op_suffix = order_point_suffix(raw_vendor_op)
                        if raw_vendor_op and not op_suffix and track:
                            issues['order_point'].append(f"{dpci_val}（{raw_vendor_op}）")

                        fact_parts = [p for p in [raw_factory_name, raw_factory_id, op_suffix] if p]
                        factory_combined = " / ".join(fact_parts)
                        
                        w_row(ws, start_row, start_col, 1, "DPCI:", dpci_val, "", "Style:", style_val,
                              'lbl_l', 'dat_in', 'dat_in', 'lbl_in', 'dat_r')
                        w_row(ws, start_row, start_col, 2, "UPC#:", upc_val, "", "PID:", pid_val,
                              'lbl_l', 'dat_in', 'dat_in', 'lbl_in', 'dat_r')
                        w_row(ws, start_row, start_col, 3, "TCIN:", "", "", "SP:", "☐",
                              'lbl_l', 'dat_in', 'dat_in', 'lbl_in', 'dat_r')
                        
                        ws.data_validation(start_row + 3, start_col + 4, start_row + 3, start_col + 4,
                                           {'validate': 'list', 'source': ['☑', '☐']})
                        ws.write(start_row + 4, start_col, "Description:", fmt['lbl_l'])
                        ws.merge_range(start_row + 4, start_col + 1, start_row + 4, start_col + 4, desc_val, fmt['dat_r'])
                        
                        w_row(ws, start_row, start_col, 5, cost_label, fca_val, "", "RETAIL:", retail_val,
                              'lbl_l', 'dat_in_curr', 'dat_in', 'lbl_in', 'dat_r_curr')
                        w_row(ws, start_row, start_col, 6, "Packaging:", pack_val, "", "Red Seal:", "",
                              'lbl_l', 'dat_in', 'dat_in', 'lbl_in', 'dat_r')
                        w_row(ws, start_row, start_col, 7, "HS NO:", hs_val, "", "Casepack:", pack_str,
                              'lbl_l', 'dat_in', 'dat_in', 'lbl_in', 'dat_r')
                        w_row(ws, start_row, start_col, 8, "Material:", mat_val, "", "QTY:", qty_val,
                              'lbl_l', 'dat_in', 'dat_in', 'lbl_in', 'dat_r')
                        
                        ws.write(start_row + 9, start_col, "Remark:", fmt['lbl_l'])
                        ws.merge_range(start_row + 9, start_col + 1, start_row + 9, start_col + 4, "", fmt['dat_r'])
                        ws.merge_range(start_row + 10, start_col, start_row + 10, start_col + 4, f"Factory: {factory_combined}", fmt['fact_merge'])

                        if temp_dir:
                            img_path = None
                            search_names = []
                            for val in [dpci_val, pid_val]:
                                if val:
                                    v_lower = val.lower()
                                    safe_v = "".join(x for x in val if x.isalnum() or x in "-_").lower()
                                    search_names.extend([f"{v_lower}.jpg", f"{v_lower}.png", f"{safe_v}.jpg", f"{safe_v}.png", f"{v_lower}.jpeg", f"{safe_v}.jpeg"])
                            
                            search_names = set(search_names) 
                            
                            for root, dirs, files in os.walk(temp_dir):
                                for file in files:
                                    if file.lower() in search_names:
                                        img_path = os.path.join(root, file)
                                        break
                                if img_path: break
                                
                            if img_path:
                                try:
                                    with Image.open(img_path) as img:
                                        img_w, img_h = img.size
                                    
                                    MAX_BOX_WIDTH = 510
                                    MAX_BOX_HEIGHT = 300
                                    
                                    scale_w = MAX_BOX_WIDTH / img_w
                                    scale_h = MAX_BOX_HEIGHT / img_h
                                    
                                    final_scale = min(scale_w, scale_h)
                                    
                                    final_w = img_w * final_scale
                                    final_h = img_h * final_scale
                                    
                                    x_offset = int((530 - final_w) / 2)
                                    y_offset = int((312 - final_h) / 2)
                                    
                                    ws.insert_image(start_row, start_col, img_path, 
                                                    {'x_scale': final_scale, 'y_scale': final_scale, 
                                                     'x_offset': max(x_offset, 5), 'y_offset': max(y_offset, 5)})
                                except Exception as e:
                                    if track:
                                        issues['no_image'].append(dpci_val or pid_val or '(無 DPCI)')
                            elif track and uploaded_img is not None:
                                issues['no_image'].append(dpci_val or pid_val or '(無 DPCI)')
                        
                        item_index += 1
                    except Exception as e:
                        if track:
                            issues['card_error'].append(f"{row.get('DPCI', '(無 DPCI)')}：{e}")
                        item_index += 1
                        continue
                
                if page_breaks:
                    ws.set_h_pagebreaks(page_breaks)

            # ========================================================
            # 🚀 【自動生成頁籤流程】
            # ========================================================
            factory_count = len(df_filtered['RawFactoryName'].unique())
            
            ws_master = workbook.add_worksheet('Master Sheet')
            draw_cards_on_sheet(ws_master, df_filtered, track=True)
            
            ws_master.write(1, 15, "Factory#:", fmt['hdr_lbl'])
            ws_master.write(1, 16, factory_count, fmt['hdr_input'])
            ws_master.write(2, 15, "Item#:", fmt['hdr_lbl'])
            ws_master.write(2, 16, len(df_filtered), fmt['hdr_input'])
            
            used_sheet_names = set(['Master Sheet'])
            
            for factory, group_df in df_filtered.groupby('RawFactoryName'):
                clean_name = str(factory).strip()
                if not clean_name: clean_name = "Unknown Factory"
                clean_name = re.sub(r'[\\/*?:\[\]]', '_', clean_name)[:28]
                
                final_name = clean_name
                counter = 1
                while final_name.lower() in [s.lower() for s in used_sheet_names]:
                    final_name = f"{clean_name}_{counter}"
                    counter += 1
                
                used_sheet_names.add(final_name)
                
                ws_fact = workbook.add_worksheet(final_name)
                draw_cards_on_sheet(ws_fact, group_df)

            workbook.close()
            
            st.success(f"排版完成！總表共包含 {len(df_filtered)} 筆商品，並已自動為您拆分成 {factory_count} 個工廠專屬頁籤。")
            # --- 資料提醒：以下狀況不會中斷生成，但交付前請確認 ---
            if issues['card_error']:
                st.error("❌ 以下商品排版時發生錯誤，卡片內容可能不完整，請檢查原始資料：\n\n"
                         + "\n".join(f"- {x}" for x in issues['card_error']))
            if issues['fob']:
                st.warning(f"⚠️ 以下 {len(issues['fob'])} 筆商品沒有 FCA 成本，已改用 FOB 成本填入 "
                           "(卡片標籤已自動改為「FOB $:」)：\n\n"
                           + "、".join(issues['fob']))
            if issues['qty_fraction']:
                st.warning("⚠️ 以下商品的 QTY 在資料檔中不是整數，卡片上已四捨五入，請回頭確認 master：\n\n"
                           + "\n".join(f"- {x}" for x in issues['qty_fraction']))
            if issues['order_point']:
                st.warning("⚠️ 以下商品的 Import Vendor Order Point 無法解析出結尾號碼，Factory 列會少一段：\n\n"
                           + "\n".join(f"- {x}" for x in issues['order_point']))
            if issues['no_image']:
                st.warning(f"⚠️ 以下 {len(issues['no_image'])} 筆商品在 SPK 檔中找不到對應縮圖 (以 DPCI / PID 比對)：\n\n"
                           + "、".join(issues['no_image']))
            st.download_button(
                label="📥 點此下載最新 Program Sheet",
                data=output.getvalue(),
                file_name=re.sub(r'[\\/:*?"<>|]', '_', program_name) + " Program Sheet.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
