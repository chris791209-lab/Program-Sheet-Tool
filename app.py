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
            
            # 準備 Excel 檔案
            output = io.BytesIO()
            workbook = xlsxwriter.Workbook(output, {'in_memory': True})
            worksheet = workbook.add_worksheet('Program sheet')
            
            # ==========================================
            # 🎨 【排版美化區】: 加入框線、底色與對齊設定
            # ==========================================
            # 資料儲存格：細框線、自動換行、垂直置中
            cell_format = workbook.add_format({
                'align': 'left', 
                'valign': 'vcenter', 
                'text_wrap': True,
                'border': 1 
            })
            
            # 標籤儲存格：粗體、淺灰底色、細框線、靠右對齊
            label_format = workbook.add_format({
                'bold': True, 
                'align': 'right', 
                'valign': 'vcenter',
                'border': 1,
                'bg_color': '#F2F2F2' 
            })
            
            # 設定大標題
            title_format = workbook.add_format({'bold': True, 'font_size': 14})
            worksheet.write(0, 0, "2025 Program Sheet Auto-Generated", title_format)

            # 動態設定欄寬 (設定 3 個區塊，每個區塊 6 欄)
            for i in range(3):
                base = i * 6
                worksheet.set_column(base, base, 13)       # 標籤 1 (如 DPCI)
                worksheet.set_column(base + 1, base + 1, 22) # 資料 1
                worksheet.set_column(base + 2, base + 2, 2)  # 卡片內部分隔小縫隙
                worksheet.set_column(base + 3, base + 3, 13) # 標籤 2 (如 Style)
                worksheet.set_column(base + 4, base + 4, 22) # 資料 2
                worksheet.set_column(base + 5, base + 5, 4)  # 卡片與卡片之間的大分隔
            
            item_index = 0
            for index, row in df.iterrows():
                try:
                    block_row = item_index // 3
                    block_col = item_index % 3
                    
                    # 基準點座標 (加入間距，每張卡片寬 6 欄，高 12 列)
                    start_row = 2 + (block_row * 12)
                    start_col = block_col * 6
                    
                    # 統一設定這張卡片的每一列高度為 20，讓畫面不擁擠
                    for r in range(start_row, start_row + 10):
                        worksheet.set_row(r, 20)
                    
                    dpci_val = str(row.get('DPCI', '')).strip()
                    if dpci_val == 'nan': dpci_val = ''
                    
                    # 1. DPCI & 2. Style (第 0 列)
                    worksheet.write(start_row, start_col, "DPCI:", label_format)
                    worksheet.write(start_row, start_col + 1, dpci_val, cell_format)
                    worksheet.write(start_row, start_col + 3, "Style:", label_format)
                    worksheet.write(start_row, start_col + 4, str(row.get('Manufacturer Style # *', '')).replace('nan',''), cell_format)
                    
                    # 12. UPC# & 13. TCIN (第 1, 2 列)
                    worksheet.write(start_row + 1, start_col, "UPC#:", label_format)
                    worksheet.write(start_row + 1, start_col + 1, str(row.get('Barcode', '')).replace('nan',''), cell_format)
                    worksheet.write(start_row + 2, start_col, "TCIN:", label_format)
                    worksheet.write(start_row + 2, start_col + 1, "", cell_format)
                    
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
                    worksheet.write(start_row + 5, start_col + 4, "", cell_format)
                    
                    # 7. HS NO & 8. Casepack (第 6 列)
                    worksheet.write(start_row + 6, start_col, "HS NO:", label_format)
                    worksheet.write(start_row + 6, start_col + 1, str(row.get('HTS Code', '')).replace('nan',''), cell_format)
                    worksheet.write(start_row + 6, start_col + 3, "Casepack:", label_format)
                    casepack = str(row.get('Case Unit Quantity', '')).replace('nan','')
                    innerpack = str(row.get('Inner Pack Unit Quantity', '')).replace('nan','')
                    worksheet.write(start_row + 6, start_col + 4, f"{casepack} / {innerpack}" if casepack else "", cell_format)
                    
                    # 9. Material & 11. QTY (第 7 列)
                    worksheet.write(start_row + 7, start_col, "Material:", label_format)
                    worksheet.write(start_row + 7, start_col + 1, str(row.get('Primary Raw Material Type', '')).replace('nan',''), cell_format)
                    worksheet.write(start_row + 7, start_col + 3, "QTY:", label_format)
                    worksheet.write(start_row + 7, start_col + 4, "", cell_format)
                    
                    # 15. Remark (第 8 列)
                    worksheet.write(start_row + 8, start_col, "Remark:", label_format)
                    worksheet.write(start_row + 8, start_col + 1, "", cell_format)
                    worksheet.write(start_row + 8, start_col + 3, "", cell_format) # 空白補齊右側框線
                    worksheet.write(start_row + 8, start_col + 4, "", cell_format) # 空白補齊右側框線
                    
                    # 10. Factory (第 9 列)
                    worksheet.write(start_row + 9, start_col, "Factory:", label_format)
                    worksheet.write(start_row + 9, start_col + 1, str(row.get('Factory Name', '')).replace('nan',''), cell_format)
                    worksheet.write(start_row + 9, start_col + 3, "", cell_format) # 空白補齊右側框線
                    worksheet.write(start_row + 9, start_col + 4, "", cell_format) # 空白補齊右側框線

                    # --- 匯入圖片邏輯 ---
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
                            # 插入圖片 (放在 Style 的下方位置)
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
