# 監造報表自動生成系統(photo_report_system)

結合影像處理與 Word 自動排版，建立輕量化工具。可透過行動裝置上傳現場拍攝的照片後，系統裡會自動判定拍攝時間並進行分類，減少手動調整照片的步驟。

# 實作重點
- **多格式相容機制及自動偵測方向**：採用 pillow_heif 套件，支援 iPhone 預設的 HEIC 高效率影像格式 與傳統 JPG/PNG，解決跨平台手機照片上傳失敗的痛點及 ImageOps.exif_transpose 自動偵測並修正手機直拍/橫拍的 EXIF 旋轉標籤，確保照片在 Word 中方向一致。
- **EXIF 偵測**：讀取照片底層的 EXIF 資訊，從中取出拍攝日期。若無法抓取 EXIF 資訊則啟動另一個機制，自動以正規表達式解析「檔案名稱」中的日期數字。
- **排版不拆頁**：操作 Word XML（OxmlElement），寫入 cantSplit（禁止列跨頁斷開）與 exact 固定高度，確保表格不會因排排版因素拆頁。
- **畫質處理**：基於 Pillow 影像演算法，將影像重新封裝為 300 DPI 高解析度，確保輸出的 Word 在列印時不會出現網頁模糊或馬賽克。

# 使用工具
- Python 3.x (python-docx, Pillow, pillow-heif, re, io)
- Streamlit Framework (Web UI / Multi-file Uploader / Component Layout)
- GitHub (Version Control)
- Streamlit Community Cloud (Serverless Deployment / CI/CD)

# 專案結構
- requirements.txt：雲端環境依賴套件設定檔，精簡配置 streamlit、python-docx、Pillow 與 pillow-heif 共 4 行核心套件，確保雲端伺服器環境完全一致。
- word.py：相片中介資料解析、影像壓縮裁切、Word XML 表格精細排版及多工互動式網頁的一體化主要程式碼。

執行成果
<img width="883" height="847" alt="實作照片" src="https://github.com/user-attachments/assets/980efeeb-53da-4108-9057-b914e8975802" />

# Streamlit 網頁測試網址
- https://usesoeasy.streamlit.app/
