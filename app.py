import streamlit as st
import pandas as pd
import io
import os
import xlsxwriter

# --- 網頁介面設定 ---
st.set_page_config(page_title="Program Sheet 生成器", layout="centered")
st.title("🚀 Program Sheet 自動生成器")
st.markdown("只需上傳 Data.csv，系統將瞬間為您排版並匯入圖片！")

# 讓使用者上傳 CSV 檔案
uploaded_file = st.file_uploader("請上傳 Data.csv", type=["csv"])
# 讓使用者輸入本機的圖片資料夾路徑
image_folder = st.text_input("請輸入本機圖片資料夾路徑 (例如: C:/Users/Chris/Desktop/Images)", "")

if uploaded_file is not None:
    st.success("資料檔上傳成功！")
    
    if st.button("生成 Program Sheet"):
        with st.spinner("正在為您進行排版與匯入圖片，請稍候..."):
            
            # 1. 讀取資料並清除 DPCI 為空的列
            df = pd.read_csv(uploaded_file, skiprows=2) # 假設前兩行是標題說明，從第3行開始讀
            if 'DPCI' in df.columns:
                df = df.dropna(subset=['DPCI'])
            
            # 2. 準備一個虛擬的 Excel 檔案放在記憶體中
            output = io.BytesIO()
            workbook = xlsxwriter.Workbook(output, {'in_memory': True})
            worksheet = workbook.add_worksheet('Program sheet')
            
            # 設定基本格式
            cell_format = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'text_wrap': True})
            label_format = workbook.add_format({'bold': True, 'align': 'right'})
            
            # 設定欄寬 (A=0, B=1, C=2... 以此類推)
            worksheet.set_column('A:A', 12)
            worksheet.set_column('B:B', 20)
            worksheet.set_column('D:D', 12)
            worksheet.set_column('E:E', 20)
            worksheet.set_column('G:G', 12)
            worksheet.set_column('H:H', 20)
            
            # 3. 核心排版邏輯 (完全移植您原本的 3xN 網格邏輯)
            item_index = 0
            for index, row in df.iterrows():
                try:
                    block_row = item_index // 3  # 第幾列
                    block_col = item_index % 3   # 左中右
                    
                    # 基準點座標 (Python 的 XlsxWriter 是從 0 開始算，所以原來的 3 變成 2，1 變成 0)
                    start_row = 2 + (block_row * 11)
                    start_col = 0 + (block_col * 6)
                    
                    # --- 寫入標籤與資料 ---
                    # DPCI
                    worksheet.write(start_row, start_col, "DPCI:", label_format)
                    worksheet.write(start_row, start_col + 1, str(row.get('DPCI', '')), cell_format)
                    
                    # Style
                    worksheet.write(start_row, start_col + 3, "Style:", label_format)
                    worksheet.write(start_row, start_col + 4, str(row.get('Manufacturer Style # *', '')), cell_format)
                    
                    # Description
                    worksheet.write(start_row + 3, start_col, "Description:", label_format)
                    worksheet.write(start_row + 3, start_col + 1, str(row.get('Vendor Product Description *', '')), cell_format)
                    
                    # --- 匯入圖片 ---
                    # 假設圖片檔名與 DPCI 完全相同，格式為 .jpg
                    if image_folder:
                        dpci_val = str(row.get('DPCI', '')).strip()
                        img_path = os.path.join(image_folder, f"{dpci_val}.jpg")
                        
                        if os.path.exists(img_path):
                            # 將圖片插入在 Style 下方的位置 (可利用 x_scale, y_scale 調整縮放比例)
                            worksheet.insert_image(start_row + 2, start_col + 4, img_path, 
                                                   {'x_scale': 0.15, 'y_scale': 0.15, 'x_offset': 5, 'y_offset': 5})
                    
                    item_index += 1
                except Exception as e:
                    continue # 如果該筆資料有問題就跳過
                    
            workbook.close()
            
            # 4. 產生下載按鈕
            st.success(f"排版完成！共處理 {item_index} 筆商品資料。")
            st.download_button(
                label="📥 點此下載最新 Program Sheet",
                data=output.getvalue(),
                file_name="Program_Sheet_Auto.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )