# 隱私與公開發布注意事項

本系統會處理學生姓名、學號、Email、作業檔案、老師評語與訪視資訊。正式使用時，資料會存放在部署老師自己的 Google Drive 與 Google Sheets。

公開到 GitHub 前請確認：

- `Code.gs` 沒有真實 `ROOT_FOLDER_ID`。
- `Code.gs` 沒有真實 `SPREADSHEET_ID`。
- `Code.gs` 沒有老師或學生的 Email。
- 沒有真實學生姓名、學號、機構地址或實習資料。
- 沒有正式部署網址。
- 沒有 `.clasp.json`，避免公開 Apps Script 專案 ID。

建議公開的是本模板版，不是已經部署使用中的正式版。
