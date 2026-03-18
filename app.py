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
            # 🖨️ 【列印與分頁設定】
            # ========================================================
            worksheet.set_landscape()
            worksheet.set_margins(left=0.3, right=0.3, top=0.4, bottom=0.4)
            worksheet.fit_to_pages(1, 0)
            
            # ========================================================
            # 🎨 【排版美化設定區】
            # ========================================================
            # 1. 一般資料欄位 (取消 border 屬性，消除細框線)
            cell_format = workbook.add_format({
                'font_name': 'Arial',
                'font_size': 10,
                'align': 'left', 
                'valign': 'vcenter', 
                'text_wrap': True
                # 已移除 'border': 1
            })
            
            # 2. 一般標籤欄位 (使用粗外框線 border: 2)
            label_format = workbook.add_format({
                'font_name': 'Arial',
                'font_size': 10,
                'bold': True, 
                'align': 'right', 
                'valign': 'vcenter',
                'border': 2,              # 2=中粗框線, 若想要更粗可改為 5
                'bg_color': '#E7E6E6'
            })

            # 3. 特殊紅色標籤 (使用粗外框線 border: 2)
            red_label_format = workbook.add_format({
                'font_name': 'Arial',
                'font_size': 10,
                'bold': True, 
                'font_color': '#FF0000',
                'align': 'right', 
                'valign': 'vcenter',
                'border': 2,
                'bg_color': '#E7E6E6'
            })
            
            # 4. 頂部圖片區塊格式 (可選配加上細外框把圖片區包起來)
            img_placeholder_format = workbook.add_format({
                'border': 1,              # 幫圖片預留區畫個細框，方便視覺辨識
                'bg_color': '#FFFFFF'
            })
            
            title_format = workbook.add_format({'font_name': 'Arial', 'bold': True, 'font_size': 14})
            worksheet.write(0, 0, "2025 Program Sheet Auto-Generated", title_format)

            # 動態設定欄寬
            for i in range(3):
                base = i * 6
                worksheet.set_column(base, base, 13)       
                worksheet.set_column(base + 1, base + 1, 22) 
                worksheet.set_column(base + 2, base + 2, 2)  
                worksheet.set_column(base + 3, base + 3, 13) 
                worksheet.set_column(base + 4, base + 4, 22) 
                worksheet.set_column(base + 5, base + 5, 4)  
            
            item_index = 0
            page_breaks = [] 
            
            for index, row in df.iterrows():
                try:
                    block_row = item_index // 3
                    block_col = item_index % 3
                    
                    # 卡片佔 13 列 (1列圖片 + 10列資料 + 2列空白)
                    start_row = 2 + (block_row * 13)
                    start_col = block_col * 6
                    
                    # ----------------------------------------------------
                    # 🖼️ 建立頂部圖片區 (第 0 列)
                    # ----------------------------------------------------
                    # 強制設定第一列高度為 234 (約 312 像素)
                    worksheet.set_row(start_row, 234)
                    
                    # 合併第 0 欄到第 4 欄作為圖片的放置區域
                    worksheet.merge_range(start_row, start_col, start_row, start_col + 4, "", img_placeholder_format)
                    
                    # ----------------------------------------------------
                    # 📝 建立下方資料區 (從第 1 列開始寫入)
                    # ----------------------------------------------------
                    # 統一設定這張卡片下方資料列的高度為 22
                    for r in range(start_row + 1, start_row + 11):
                        worksheet.set_row(r, 22)
                    
                    if block_row > 0 and block_row % 2 == 0 and block_col == 0:
                        page_breaks.append(start_row - 1)
                    
                    dpci_val = str(row.get('DPCI', '')).strip()
                    if dpci_val == 'nan': dpci_val = ''
                    
                    # DPCI & Style (第 1 列)
                    worksheet.write(start_row + 1, start_col, "DPCI:", label_format)
                    worksheet.write(start_row + 1, start_col + 1, dpci_val, cell_format)
                    worksheet.write(start_row + 1, start_col + 3, "Style:", label_format)
                    worksheet.write(start_row + 1, start_col + 4, str(row.get('Manufacturer Style # *', '')).replace('nan',''), cell_format)
                    
                    # UPC# & TCIN (第 2, 3 列)
                    worksheet.write(start_row + 2, start_col, "UPC#:", label_format)
                    worksheet.write(start_row + 2, start_col + 1, str(row.get('Barcode', '')).replace('nan',''), cell_format)
                    worksheet.write(start_row + 3, start_col, "TCIN:", label_format)
                    worksheet.write(start_row + 3, start_col + 1, "", cell_format)
                    
                    # Description (第 4 列)
                    worksheet.write(start_row + 4, start_col, "Description:", label_format)
                    worksheet.write(start_row + 4, start_col + 1, str(row.get('Vendor Product Description *', '')).replace('nan',''), cell_format)
                    
                    # FCA & RETAIL (第 5 列)
                    worksheet.write(start_row + 5, start_col, "FCA $:", red_label_format)
                    worksheet.write(start_row + 5, start_col + 1, str(row.get('FCA Factory City Unit Cost', '')).replace('nan',''), cell_format)
                    worksheet.write(start_row + 5, start_col + 3, "RETAIL:", red_label_format)
                    worksheet.write(start_row + 5, start_col + 4, str(row.get('Suggested Unit Retail', '')).replace('nan',''), cell_format)
                    
                    # Packaging & Red Seal (第 6 列)
                    worksheet.write(start_row + 6, start_col, "Packaging:", label_format)
                    worksheet.write(start_row + 6, start_col + 1, str(row.get('Retail Packaging Format (1) *', '')).replace('nan',''), cell_format)
                    worksheet.write(start_row + 6, start_col + 3, "Red Seal:", label_format)
                    worksheet.write(start_row + 6, start_col + 4, "", cell_format)
                    
                    # HS NO & Casepack (第 7 列)
                    worksheet.write(start_row + 7, start_col, "HS NO:", label_format)
                    worksheet.write(start_row + 7, start_col + 1, str(row.get('HTS Code', '')).replace('nan',''), cell_format)
                    worksheet.write(start_row + 7, start_col + 3, "Casepack:", label_format)
                    casepack = str(row.get('Case Unit Quantity', '')).replace('nan','')
                    innerpack = str(row.get('Inner Pack Unit Quantity', '')).replace('nan','')
                    worksheet.write(start_row + 7, start_col + 4, f"{casepack} / {innerpack}" if casepack else "", cell_format)
                    
                    # Material & QTY (第 8 列)
                    worksheet.write(start_row + 8, start_col, "Material:", label_format)
                    worksheet.write(start_row + 8, start_col + 1, str(row.get('Primary Raw Material Type', '')).replace('nan',''), cell_format)
                    worksheet.write(start_row + 8, start_col + 3, "QTY:", red_label_format)
                    worksheet.write(start_row + 8, start_col + 4, "", cell_format)
                    
                    # Remark (第 9 列)
                    worksheet.write(start_row + 9, start_col, "Remark:", label_format)
                    worksheet.write(start_row + 9, start_col + 1, "", cell_format)
                    worksheet.write(start_row + 9, start_col + 3, "", cell_format) 
                    worksheet.write(start_row + 9, start_col + 4, "", cell_format) 
                    
                    # Factory (第 10 列)
                    worksheet.write(start_row + 10, start_col, "Factory:", label_format)
                    worksheet.write(start_row + 10, start_col + 1, str(row.get('Factory Name', '')).replace('nan',''), cell_format)
                    worksheet.write(start_row + 10, start_col + 3, "", cell_format) 
                    worksheet.write(start_row + 10, start_col + 4, "", cell_format) 

                    # --- 匯入圖片 (位置改至最上方的圖片區) ---
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
                            # 插入圖片 (放在 start_row 的圖片區內)
                            # 如果您發現上傳的圖片太大或太小，可以微調 x_scale 與 y_scale
                            # x_offset 與 y_offset 負責將圖片從格子的左上角稍微往中間推
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
