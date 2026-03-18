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
                df = pd.read_excel(uploaded_file, sheet_name="Data", skiprows=2, engine="openpyxl")
            except Exception as e:
                st.error(f"讀取 Excel 失敗，請確認檔案內是否有名為 'Data' 的工作表頁籤。錯誤訊息: {e}")
                st.stop()
                
            if 'DPCI' in df.columns:
                df = df.dropna(subset=['DPCI'])
            
            temp_dir = None
            if uploaded_zip is not None:
                temp_dir = tempfile.mkdtemp()
                with zipfile.ZipFile(uploaded_zip, 'r') as zip_ref:
                    zip_ref.extractall(temp_dir)
            
            output = io.BytesIO()
            workbook = xlsxwriter.Workbook(output, {'in_memory': True})
            worksheet = workbook.add_worksheet('Program sheet')
            
            # ========================================================
            # 🖨️ 【列印與分頁設定 (Page Setup)】: 確保一頁剛好塞滿 6 張卡片
            # ========================================================
            worksheet.set_landscape() # 設定為橫向列印 (3張卡片橫排通常需要橫向)
            worksheet.set_margins(left=0.3, right=0.3, top=0.4, bottom=0.4) # 縮小邊距
            worksheet.fit_to_pages(1, 0) # 寬度強制縮放為 1 頁，長度無限延伸
            
            # ========================================================
            # 🎨 【排版美化設定區】: 完全復刻範本的字體、顏色、粗細框線
            # ========================================================
            # 1. 一般資料欄位 (白底、左對齊、細框線)
            cell_format = workbook.add_format({
                'font_name': 'Arial',     # 可改為 'Calibri', '微軟正黑體' 等
                'font_size': 10,
                'align': 'left', 
                'valign': 'vcenter', 
                'text_wrap': True,
                'border': 1               # 1=細框線, 2=粗框線
            })
            
            # 2. 一般標籤欄位 (灰底、粗體、右對齊、細框線)
            label_format = workbook.add_format({
                'font_name': 'Arial',
                'font_size': 10,
                'bold': True, 
                'align': 'right', 
                'valign': 'vcenter',
                'border': 1,
                'bg_color': '#E7E6E6'     # 您可以更換為範本確切的 Hex 色碼
            })

            # 3. 特殊標籤 (例如紅字標籤 Red Seal)
            red_label_format = workbook.add_format({
                'font_name': 'Arial',
                'font_size': 10,
                'bold': True, 
                'font_color': '#FF0000',  # 紅色字體
                'align': 'right', 
                'valign': 'vcenter',
                'border': 1,
                'bg_color': '#E7E6E6'
            })
            
            # 設定大標題
            title_format = workbook.add_format({'font_name': 'Arial', 'bold': True, 'font_size': 14})
            worksheet.write(0, 0, "2025 Program Sheet Auto-Generated", title_format)

            # 動態設定欄寬 (3 個區塊，每個區塊 5 欄 + 1 欄間距 = 6 欄)
            for i in range(3):
                base = i * 6
                worksheet.set_column(base, base, 13)       # 標籤 1 (如 DPCI)
                worksheet.set_column(base + 1, base + 1, 22) # 資料 1
                worksheet.set_column(base + 2, base + 2, 2)  # 內部小間隔
                worksheet.set_column(base + 3, base + 3, 13) # 標籤 2 (如 Style)
                worksheet.set_column(base + 4, base + 4, 22) # 資料 2
                worksheet.set_column(base + 5, base + 5, 4)  # 卡片與卡片之間的外部大分隔
            
            item_index = 0
            page_breaks = [] # 用來記錄分頁符號的位置
            
            for index, row in df.iterrows():
                try:
                    block_row = item_index // 3
                    block_col = item_index % 3
                    
                    # 基準點座標 (每張卡片高 10 列，加上 2 列空白間距，總共佔 12 列)
                    start_row = 2 + (block_row * 12)
                    start_col = block_col * 6
                    
                    # 統一設定這張卡片的列高為 22，提供充裕的空間
                    for r in range(start_row, start_row + 10):
                        worksheet.set_row(r, 22)
                    
                    # --- 每隔 2 列卡片 (即 6 張商品)，插入一個強制的水平分頁符號 ---
                    if block_row > 0 and block_row % 2 == 0 and block_col == 0:
                        page_breaks.append(start_row - 1)
                    
                    dpci_val = str(row.get('DPCI', '')).strip()
                    if dpci_val == 'nan': dpci_val = ''
                    
                    # ==========================================
                    # 【資料寫入區】: 套用上方設定好的 format
                    # ==========================================
                    worksheet.write(start_row, start_col, "DPCI:", label_format)
                    worksheet.write(start_row, start_col + 1, dpci_val, cell_format)
                    worksheet.write(start_row, start_col + 3, "Style:", label_format)
                    worksheet.write(start_row, start_col + 4, str(row.get('Manufacturer Style # *', '')).replace('nan',''), cell_format)
                    
                    worksheet.write(start_row + 1, start_col, "UPC#:", label_format)
                    worksheet.write(start_row + 1, start_col + 1, str(row.get('Barcode', '')).replace('nan',''), cell_format)
                    worksheet.write(start_row + 2, start_col, "TCIN:", label_format)
                    worksheet.write(start_row + 2, start_col + 1, "", cell_format)
                    
                    worksheet.write(start_row + 3, start_col, "Description:", label_format)
                    worksheet.write(start_row + 3, start_col + 1, str(row.get('Vendor Product Description *', '')).replace('nan',''), cell_format)
                    
                    worksheet.write(start_row + 4, start_col, "FCA $:", label_format)
                    worksheet.write(start_row + 4, start_col + 1, str(row.get('FCA Factory City Unit Cost', '')).replace('nan',''), cell_format)
                    worksheet.write(start_row + 4, start_col + 3, "RETAIL:", label_format)
                    worksheet.write(start_row + 4, start_col + 4, str(row.get('Suggested Unit Retail', '')).replace('nan',''), cell_format)
                    
                    worksheet.write(start_row + 5, start_col, "Packaging:", label_format)
                    worksheet.write(start_row + 5, start_col + 1, str(row.get('Retail Packaging Format (1) *', '')).replace('nan',''), cell_format)
                    # 使用紅色字體標籤格式 (red_label_format)
                    worksheet.write(start_row + 5, start_col + 3, "Red Seal:", red_label_format)
                    worksheet.write(start_row + 5, start_col + 4, "", cell_format)
                    
                    worksheet.write(start_row + 6, start_col, "HS NO:", label_format)
                    worksheet.write(start_row + 6, start_col + 1, str(row.get('HTS Code', '')).replace('nan',''), cell_format)
                    worksheet.write(start_row + 6, start_col + 3, "Casepack:", label_format)
                    casepack = str(row.get('Case Unit Quantity', '')).replace('nan','')
                    innerpack = str(row.get('Inner Pack Unit Quantity', '')).replace('nan','')
                    worksheet.write(start_row + 6, start_col + 4, f"{casepack} / {innerpack}" if casepack else "", cell_format)
                    
                    worksheet.write(start_row + 7, start_col, "Material:", label_format)
                    worksheet.write(start_row + 7, start_col + 1, str(row.get('Primary Raw Material Type', '')).replace('nan',''), cell_format)
                    worksheet.write(start_row + 7, start_col + 3, "QTY:", label_format)
                    worksheet.write(start_row + 7, start_col + 4, "", cell_format)
                    
                    worksheet.write(start_row + 8, start_col, "Remark:", label_format)
                    worksheet.write(start_row + 8, start_col + 1, "", cell_format)
                    worksheet.write(start_row + 8, start_col + 3, "", cell_format) 
                    worksheet.write(start_row + 8, start_col + 4, "", cell_format) 
                    
                    worksheet.write(start_row + 9, start_col, "Factory:", label_format)
                    worksheet.write(start_row + 9, start_col + 1, str(row.get('Factory Name', '')).replace('nan',''), cell_format)
                    worksheet.write(start_row + 9, start_col + 3, "", cell_format) 
                    worksheet.write(start_row + 9, start_col + 4, "", cell_format) 

                    # --- 匯入圖片 ---
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
                            worksheet.insert_image(start_row + 1, start_col + 4, img_path, 
                                                   {'x_scale': 0.16, 'y_scale': 0.16, 'x_offset': 5, 'y_offset': 5})
                    
                    item_index += 1
                except Exception as e:
                    st.warning(f"處理第 {item_index+1} 筆資料時發生錯誤: {e}")
                    continue 
            
            # 寫入所有的水平分頁符號
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
