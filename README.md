# Internship Submission Portal

一套以 Google Apps Script 建置的實習作業上傳、教師批閱、退件通知與訪視安排系統。系所老師可以複製此專案到自己的 Google 帳號，使用自己的 Google Drive 與 Google Sheets 保存資料。

## 功能

- 學生使用 Google 帳號登入後上傳作業。
- 學生只能看到自己的上傳紀錄與訪視回報。
- 老師可以查看全班繳交狀態、退件、填寫退件原因、勾選已完成。
- 退件時可寄送 Email 通知學生。
- 老師可上傳批閱後檔案給學生下載。
- 公告欄可由老師新增、修改與置頂。
- 名單管理支援貼上 Excel/試算表資料後匯入。
- 學生可回報訪視日期、時間、地址、聯絡人與備註。
- 老師可在訪視後台核對回報資料，並加入 Google Calendar。

## 專案結構

```text
apps_script/
  Code.gs
  Index.html
  appsscript.json
docs/
  teacher-setup.md
  privacy.md
  student-manual-template.md
```

## 快速安裝

1. 建立新的 Google Apps Script 專案。
2. 將 `apps_script/Code.gs`、`apps_script/Index.html`、`apps_script/appsscript.json` 複製到 Apps Script。
3. 在 Apps Script 編輯器執行：

```javascript
runInitialSetup("teacher@example.com")
```

多位老師可用逗號分隔：

```javascript
runInitialSetup("teacher1@example.com,teacher2@example.com")
```

4. 第一次執行時，依畫面完成 Google 授權。
5. 部署為網頁應用程式：
   - 執行身分：我
   - 存取權：任何擁有 Google 帳戶的使用者
6. 將部署網址提供給學生。

## 修改系統名稱

預設名稱在 `Code.gs`：

```javascript
const DEFAULT_COURSE_TITLE = '實習作業上傳與評分系統';
```

老師可以改成自己的課程或系所名稱。

## 修改作業項目

作業項目在 `Code.gs` 的 `ASSIGNMENTS` 陣列。可依課程需求新增、刪除或改名。

## 匯入學生名單

老師登入系統後，到「名單管理」貼上含標題列的 Excel 或 Google Sheets 資料。建議欄位：

- 學號
- 姓名
- 班級
- 實習機構
- 實習機構地址
- 實習機構督導
- Google 帳號（Gmail）

若一開始沒有 Gmail，也可以先留空，讓學生第一次上傳時填寫自己的 Google 帳號。

## 隱私提醒

請勿把已含學生名單、Google Drive ID、試算表 ID、部署網址或老師個人信箱的版本公開到 GitHub。公開前請執行敏感資料掃描。

## License

MIT License
