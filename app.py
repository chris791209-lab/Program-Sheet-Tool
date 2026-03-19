import streamlit as st
import pandas as pd
import io
import os
import xlsxwriter
import zipfile
import tempfile

# --- 網頁介面設定 ---
st.set_page_config(page_title="Program Sheet 生成器", layout="centered")
st.title("🚀 Program Sheet 自動生成器")
st.markdown("只需上傳 Data 檔案 (.xlsm 或 .xlsx) 與圖片壓縮檔，系統將瞬間為您排版並匯入圖片！")

uploaded_file = st.file_uploader("1. 請上傳包含資料的 Excel 檔 (.xlsm / .xlsx)", type=["xlsm", "xlsx"])
uploaded_zip = st.file_uploader("2. (選填) 請上傳包含產品圖片的 .zip 壓縮檔", type=["zip"])

if uploaded_file is not None:
    st.success("資料檔已就緒！")
    
    if st.button("✨ 生成 Program Sheet"):
        with st.spinner("正在為您進行排版與處理圖片，請稍候..."):
            
            try:
                # 【修正 1】移除 skiprows=2。讓 Python 正確讀取第一列的標題與第二列的資料
                df = pd.read_excel(uploaded_file, sheet_name="Data", engine="openpyxl")
            except Exception as e:
                st.error(f"讀取 Excel 失敗，請確認檔案內是否有名為 'Data' 的工作表頁籤。錯誤訊息: {e}")
                st.stop()
                
            # 【修正 2】確保 DPCI (第7欄，Index=6) 有資料才保留
            if len(df.columns) >= 7:
                dpci_col_name = df.columns[6]
                df = df.dropna(subset=[dpci_col_name])
            
            temp_dir = None
            if uploaded_zip is not None:
                temp_dir = tempfile.mkdtemp()
                with zipfile.ZipFile(uploaded_zip, 'r') as zip_ref:
                    zip_ref.extractall(temp_dir)
            
            output = io.BytesIO()
            workbook = xlsxwriter.Workbook(output, {'in_memory': True})
            worksheet = workbook.add_worksheet('Program sheet')
            
            # ========================================================
            # 🖨️ 【列印與分頁設定】
            # ========================================================
            worksheet.set_landscape()
            worksheet.set_margins(left=0.3, right=0.3, top=0.4, bottom=0.4)
            worksheet.fit_to_pages(1, 0)
            
            # ========================================================
            # 🎨 【排版美化設定區】
            # ========================================================
            base_props = {
                # 【修正 3】將 label 與 red_label 的 align 從 right 改為 left
                'label': {'font_name': 'Arial', 'font_size': 10, 'bold': True, 'align': 'left', 'valign': 'vcenter', 'bg_color': '#E7E6E6'},
                'red_label': {'font_name': 'Arial', 'font_size': 10, 'bold': True, 'font_color': '#FF0000', 'align': 'left', 'valign': 'vcenter', 'bg_color': '#E7E6E6'},
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
            create_fmt('rlbl_l', 'red_label', left=2)
            create_fmt('rlbl_in', 'red_label')
            create_fmt('dat_in', 'data')
            create_fmt('dat_r', 'data', right=2)
            create_fmt('dat_b', 'data', bottom=2)
            create_fmt('dat_rb', 'data', right=2, bottom=2)

            title_format = workbook.add_format({'font_name': 'Arial', 'bold': True, 'font_size': 14})
            worksheet.write(0, 0, "2025 Program Sheet Auto-Generated", title_format)

            for i in range(3):
                base = i * 6
                worksheet.set_column(base, base, 13)       
                worksheet.set_column(base + 1, base + 1, 22) 
                worksheet.set_column(base + 2, base + 2, 2)  
                worksheet.set_column(base + 3, base + 3, 13) 
                worksheet.set_column(base + 4, base + 4, 22) 
                worksheet.set_column(base + 5, base + 5, 4)  
            
            def w_row(ws, s_row, s_col, r_off, c0, c1, c2, c3, c4, f0, f1, f2, f3, f4):
                ws.write(s_row + r_off, s_col + 0, c0, fmt[f0])
                ws.write(s_row + r_off, s_col + 1, c1, fmt[f1])
                ws.write(s_row + r_off, s_col + 2, c2, fmt[f2]) 
                ws.write(s_row + r_off, s_col + 3, c3, fmt[f3])
                ws.write(s_row + r_off, s_col + 4, c4, fmt[f4])

            # ========================================================
            # 🛠️ 【資料抓取神小幫手】: 就像 VBA 一樣，直接輸入「第幾欄」就能精準抓資料
            # ========================================================
            def get_val(row_series, col_num):
                try:
                    # 避免數字超出總欄位數量 (col_num - 1 是因為 Python index 從 0 開始)
                    if (col_num - 1) < len(row_series):
                        val = str(row_series.iloc[col_num - 1]).strip()
                        return "" if val.lower() == 'nan' else val
                    return ""
                except:
                    return ""

            item_index = 0
            page_breaks = [] 
            
            for index, row in df.iterrows():
                try:
                    block_row = item_index // 3
                    block_col = item_index % 3
                    
                    start_row = 2 + (block_row * 13)
                    start_col = block_col * 6
                    
                    worksheet.set_row(start_row, 234) 
                    worksheet.merge_range(start_row, start_col, start_row, start_col + 4, "", fmt['img_top'])
                    
                    for r in range(start_row + 1, start_row + 11):
                        worksheet.set_row(r, 22)
                    
                    if block_row > 0 and block_row % 2 == 0 and block_col == 0:
                        page_breaks.append(start_row - 1)
                    
                    # ----------------------------------------------------
                    # 🎯 【資料對應區】：完美對齊您最初 VBA 成功的欄位設定
                    # ----------------------------------------------------
                    dpci_val = get_val(row, 7)     # DPCI: 第 7 欄
                    style_val = get_val(row, 14)   # Style: 第 14 欄
                    upc_val = get_val(row, 13)     # UPC: 假設在第 13 欄 (Barcode)
                    desc_val = get_val(row, 4)     # Description: 第 4 欄
                    fca_val = get_val(row, 26)     # FCA: 第 26 欄
                    retail_val = get_val(row, 25)  # RETAIL: 第 25 欄
                    pack_val = get_val(row, 70)    # Packaging: 第 70 欄
                    hs_val = get_val(row, 56)      # HS NO: 第 56 欄
                    
                    case_q = get_val(row, 27)      # 外箱: 第 27 欄
                    inner_q = get_val(row, 32)     # 內箱: 第 32 欄
                    pack_str = f"{case_q} / {inner_q}" if (case_q or inner_q) else ""
                    
                    mat_val = get_val(row, 61)     # Material: 第 61 欄
                    
                    fact_1 = get_val(row, 100)     # Factory(外): 第 100 欄
                    fact_2 = get_val(row, 90)      # Factory(內): 第 90 欄
                    factory_val = f"{fact_1} - {fact_2}".strip(" - ")
                    
                    w_row(worksheet, start_row, start_col, 1, "DPCI:", dpci_val, "", "Style:", style_val,
                          'lbl_l', 'dat_in', 'dat_in', 'lbl_in', 'dat_r')
                    
                    w_row(worksheet, start_row, start_col, 2, "UPC#:", upc_val, "", "", "",
                          'lbl_l', 'dat_in', 'dat_in', 'dat_in', 'dat_r')
                    
                    w_row(worksheet, start_row, start_col, 3, "TCIN:", "", "", "", "",
                          'lbl_l', 'dat_in', 'dat_in', 'dat_in', 'dat_r')
                    
                    w_row(worksheet, start_row, start_col, 4, "Description:", desc_val, "", "", "",
                          'lbl_l', 'dat_in', 'dat_in', 'dat_in', 'dat_r')
                    
                    w_row(worksheet, start_row, start_col, 5, "FCA $:", fca_val, "", "RETAIL:", retail_val,
                          'rlbl_l', 'dat_in', 'dat_in', 'rlbl_in', 'dat_r')
                    
                    w_row(worksheet, start_row, start_col, 6, "Packaging:", pack_val, "", "Red Seal:", "",
                          'lbl_l', 'dat_in', 'dat_in', 'rlbl_in', 'dat_r')
                    
                    w_row(worksheet, start_row, start_col, 7, "HS NO:", hs_val, "", "Casepack:", pack_str,
                          'lbl_l', 'dat_in', 'dat_in', 'lbl_in', 'dat_r')
                    
                    w_row(worksheet, start_row, start_col, 8, "Material:", mat_val, "", "QTY:", "",
                          'lbl_l', 'dat_in', 'dat_in', 'rlbl_in', 'dat_r')
                    
                    w_row(worksheet, start_row, start_col, 9, "Remark:", "", "", "", "",
                          'lbl_l', 'dat_in', 'dat_in', 'dat_in', 'dat_r')
                    
                    w_row(worksheet, start_row, start_col, 10, "Factory:", factory_val, "", "", "",
                          'lbl_lb', 'dat_b', 'dat_b', 'dat_b', 'dat_rb')

                    if temp_dir and dpci_val:
                        img_path = None
                        for root, dirs, files in os.walk(temp_dir):
                            for file in files:
                                if file.lower() == f"{dpci_val.lower()}.jpg" or file.lower() == f"{dpci_val.lower()}.png":
                                    img_path = os.path.join(root, file)
                                    break
                            if img_path: 
                                break
                        
                        if img_path:
                            worksheet.insert_image(start_row, start_col, img_path, 
                                                   {'x_scale': 0.28, 'y_scale': 0.28, 'x_offset': 15, 'y_offset': 15})
                    
                    item_index += 1
                except Exception as e:
                    st.warning(f"處理第 {item_index+1} 筆資料時發生錯誤: {e}")
                    continue 
            
            if page_breaks:
                worksheet.set_h_pagebreaks(page_breaks)
                
            workbook.close()
            
            st.success(f"排版完成！共處理 {item_index} 筆商品資料。")
            st.download_button(
                label="📥 點此下載最新 Program Sheet",
                data=output.getvalue(),
                file_name="Program_Sheet_Auto.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
