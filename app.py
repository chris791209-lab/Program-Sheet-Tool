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

# 讓使用者上傳 Excel 與 ZIP 檔案
uploaded_file = st.file_uploader("1. 請上傳包含資料的 Excel 檔 (.xlsm / .xlsx)", type=["xlsm", "xlsx"])
uploaded_zip = st.file_uploader("2. (選填) 請上傳包含產品圖片的 .zip 壓縮檔", type=["zip"])

if uploaded_file is not None:
    st.success("資料檔已就緒！")
    
    if st.button("✨ 生成 Program Sheet"):
        with st.spinner("正在為您進行排版與處理圖片，請稍候..."):
            
            # 1. 讀取 Excel 中名為 "Data" 的頁籤，跳過前兩列標題，並清除 DPCI 為空的列
            try:
                df = pd.read_excel(uploaded_file, sheet_name="Data", skiprows=2, engine="openpyxl")
            except Exception as e:
                st.error(f"讀取 Excel 失敗，請確認檔案內是否有名為 'Data' 的工作表頁籤。錯誤訊息: {e}")
                st.stop()
                
            if 'DPCI' in df.columns:
                df = df.dropna(subset=['DPCI'])
            
            # 2. 如果有上傳 ZIP，在雲端建立一個暫存資料夾來解壓縮圖片
            temp_dir = None
            if uploaded_zip is not None:
                temp_dir = tempfile.mkdtemp() # 建立隨機暫存資料夾
                with zipfile.ZipFile(uploaded_zip, 'r') as zip_ref:
                    zip_ref.extractall(temp_dir)
            
            # 3. 準備 Excel 檔案 (XlsxWriter)
            output = io.BytesIO()
            workbook = xlsxwriter.Workbook(output, {'in_memory': True})
            worksheet = workbook.add_worksheet('Program sheet')
            
            # 設定基本格式
            cell_format = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'text_wrap': True})
            label_format = workbook.add_format({'bold': True, 'align': 'right'})
            
            # 設定欄寬 (讓圖片跟文字有足夠的空間)
            worksheet.set_column('A:A', 13) # 標籤欄
            worksheet.set_column('B:B', 22) # 資料欄1
            worksheet.set_column('D:D', 13) # 標籤欄
            worksheet.set_column('E:E', 22) # 資料欄2
            
            # 4. 核心排版邏輯
            item_index = 0
            for index, row in df.iterrows():
                try:
                    block_row = item_index // 3
                    block_col = item_index // 3 * 0 + (item_index % 3) # 修正換欄邏輯
                    
                    # 基準點座標 (Python 是從 0 開始算，所以原 VBA 的 3 變成 2，1 變成 0)
                    start_row = 2 + (block_row * 11)
                    start_col = 0 + (block_col * 5) # 依據原本的寬度為 5 欄間距
                    
                    dpci_val = str(row.get('DPCI', '')).strip()
                    if dpci_val == 'nan': dpci_val = ''
                    
                    # ========================================================
                    # 【資料寫入區】: 依據版面配置，填入標籤與 Data 表的欄位值
                    # 注意: row.get('欄位名稱', '') 中的名稱必須與 Data 頁籤的標題完全一致
                    # ========================================================
                    
                    # 1. DPCI & 2. Style (第 0 列)
                    worksheet.write(start_row, start_col, "DPCI:", label_format)
                    worksheet.write(start_row, start_col + 1, dpci_val, cell_format)
                    worksheet.write(start_row, start_col + 3, "Style:", label_format)
                    worksheet.write(start_row, start_col + 4, str(row.get('Manufacturer Style # *', '')).replace('nan',''), cell_format)
                    
                    # 12. UPC# & 13. TCIN (第 1, 2 列)
                    worksheet.write(start_row + 1, start_col, "UPC#:", label_format)
                    worksheet.write(start_row + 1, start_col + 1, str(row.get('Barcode', '')).replace('nan',''), cell_format)
                    worksheet.write(start_row + 2, start_col, "TCIN:", label_format)
                    worksheet.write(start_row + 2, start_col + 1, "", cell_format) # 若無 TCIN 欄位先留白
                    
                    # 3. Description (第 3 列)
                    worksheet.write(start_row + 3, start_col, "Description:", label_format)
                    worksheet.write(start_row + 3, start_col + 1, str(row.get('Vendor Product Description *', '')).replace('nan',''), cell_format)
                    
                    # 4. FCA & 5. RETAIL (第 4 列)
                    worksheet.write(start_row + 4, start_col, "FCA $:", label_format)
                    worksheet.write(start_row + 4, start_col + 1, str(row.get('FCA Factory City Unit Cost', '')).replace('nan',''), cell_format)
                    worksheet.write(start_row + 4, start_col + 3, "RETAIL:", label_format)
                    worksheet.write(start_row + 4, start_col + 4, str(row.get('Suggested Unit Retail', '')).replace('nan',''), cell_format)
                    
                    # 6. Packaging & 14. Red Seal (第 5 列)
                    worksheet.write(start_row + 5, start_col, "Packaging:", label_format)
                    worksheet.write(start_row + 5, start_col + 1, str(row.get('Retail Packaging Format (1) *', '')).replace('nan',''), cell_format)
                    worksheet.write(start_row + 5, start_col + 3, "Red Seal:", label_format)
                    
                    # 7. HS NO & 8. Casepack (第 6 列)
                    worksheet.write(start_row + 6, start_col, "HS NO:", label_format)
                    worksheet.write(start_row + 6, start_col + 1, str(row.get('HTS Code', '')).replace('nan',''), cell_format)
                    worksheet.write(start_row + 6, start_col + 3, "Casepack:", label_format)
                    # 裝箱數合併範例 (若欄位名稱不同請自行修改字串)
                    casepack = str(row.get('Case Unit Quantity', '')).replace('nan','')
                    innerpack = str(row.get('Inner Pack Unit Quantity', '')).replace('nan','')
                    if casepack and innerpack:
                        worksheet.write(start_row + 6, start_col + 4, f"{casepack} / {innerpack}", cell_format)
                    
                    # 9. Material & 11. QTY (第 7 列)
                    worksheet.write(start_row + 7, start_col, "Material:", label_format)
                    worksheet.write(start_row + 7, start_col + 1, str(row.get('Primary Raw Material Type', '')).replace('nan',''), cell_format)
                    worksheet.write(start_row + 7, start_col + 3, "QTY:", label_format)
                    
                    # 15. Remark (第 8 列)
                    worksheet.write(start_row + 8, start_col, "Remark:", label_format)
                    
                    # 10. Factory (第 9 列)
                    worksheet.write(start_row + 9, start_col, "Factory:", label_format)
                    worksheet.write(start_row + 9, start_col + 1, str(row.get('Factory Name', '')).replace('nan',''), cell_format)

                    # --- 匯入圖片邏輯 ---
                    if temp_dir and dpci_val:
                        img_path = None
                        # 尋找檔名為 DPCI.jpg 或 DPCI.png 的圖片
                        for root, dirs, files in os.walk(temp_dir):
                            for file in files:
                                if file.lower() == f"{dpci_val.lower()}.jpg" or file.lower() == f"{dpci_val.lower()}.png":
                                    img_path = os.path.join(root, file)
                                    break
                            if img_path: 
                                break
                        
                        if img_path:
                            # 插入圖片 (放在 Style 的下方位置)
                            # 如果圖片太大或太小，請調整 x_scale 與 y_scale 的數值 (例如 0.2 或 0.1)
                            worksheet.insert_image(start_row + 1, start_col + 4, img_path, 
                                                   {'x_scale': 0.16, 'y_scale': 0.16, 'x_offset': 5, 'y_offset': 5})
                    
                    item_index += 1
                except Exception as e:
                    st.warning(f"處理第 {item_index+1} 筆資料時發生錯誤: {e}")
                    continue 
                    
            workbook.close()
            
            st.success(f"排版完成！共處理 {item_index} 筆商品資料。")
            st.download_button(
                label="📥 點此下載最新 Program Sheet",
                data=output.getvalue(),
                file_name="Program_Sheet_Auto.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
