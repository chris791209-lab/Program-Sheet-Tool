import openpyxl
from openpyxl_image_loader import SheetImageLoader
import os

# ==========================================
# ⚙️ 【參數設定區】: 執行前請依據您的 Excel 修改這裡
# ==========================================
EXCEL_FILE = 'Your_Target_Export.xlsx' # 替換成您從系統匯出的 Excel 檔名
SHEET_NAME = 'Products'                # 替換成包含圖片的工作表名稱 (依據附件應為 Products)
IMAGE_COL = 'A'                        # Thumbnail 縮圖所在的欄位 (通常是 A 欄)
DPCI_COL = 'F'                         # DPCI 或 PID 所在的欄位 (假設在 F 欄，請修改為實際字母)
START_ROW = 2                          # 資料從第幾列開始 (第 1 列通常是標題)

# ==========================================
# 🚀 【自動化萃取核心】
# ==========================================
OUTPUT_FOLDER = 'Extracted_Images'
if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)

print(f"⏳ 正在讀取 Excel 檔案: {EXCEL_FILE} ... (如果檔案很大請稍候)")
try:
    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=False)
    sheet = wb[SHEET_NAME]
except Exception as e:
    print(f"❌ 讀取失敗，請確認檔名與頁籤名稱是否正確。錯誤: {e}")
    exit()

print("🔍 正在解析圖片位置...")
image_loader = SheetImageLoader(sheet)

max_row = sheet.max_row
success_count = 0

print("🚀 開始萃取圖片並自動命名...")
for row in range(START_ROW, max_row + 1):
    image_cell = f"{IMAGE_COL}{row}"
    dpci_cell = f"{DPCI_COL}{row}"
    
    # 取得該列的 DPCI 或 PID
    dpci_value = sheet[dpci_cell].value
    
    if dpci_value:
        dpci_str = str(dpci_value).strip()
        
        # 檢查該 Thumbnail 儲存格是否有圖片
        if image_loader.image_in(image_cell):
            try:
                # 抓取圖片物件
                image = image_loader.get(image_cell)
                
                # 清洗檔名 (去除 / \ : * 等不能當檔名的字元)
                safe_filename = "".join(x for x in dpci_str if x.isalnum() or x in "-_")
                img_path = os.path.join(OUTPUT_FOLDER, f"{safe_filename}.jpg")
                
                # 轉換為 RGB 模式 (避免去背的 PNG 轉 JPG 時報錯) 並存檔
                image = image.convert('RGB')
                image.save(img_path)
                
                print(f"✅ 成功匯出: {safe_filename}.jpg")
                success_count += 1
            except Exception as e:
                print(f"⚠️ 匯出 {dpci_str} 圖片失敗: {e}")

print("="*40)
print(f"🎉 任務完成！共成功匯出 {success_count} 張圖片。")
print(f"📁 圖片已存放在專案目錄下的 '{OUTPUT_FOLDER}' 資料夾中。")