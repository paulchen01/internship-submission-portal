# 教師設定指南

## 1. 建立 Apps Script 專案

進入 Google Drive，新增 Google Apps Script 專案，貼上本專案 `apps_script` 內的三個檔案。

## 2. 初始化系統

在 Apps Script 編輯器上方選擇 `runInitialSetup`，或在函式呼叫中填入：

```javascript
runInitialSetup("teacher@example.com")
```

系統會自動建立：

- Google Drive 根資料夾
- Google Sheets 紀錄表
- 上傳紀錄表
- 公告表
- 學生名單表
- 訪視回報表

## 3. 部署網頁應用程式

選擇「部署」到「新增部署作業」，類型選「網頁應用程式」。

建議設定：

- 執行身分：我
- 存取權：任何擁有 Google 帳戶的使用者

部署後取得網址，提供給學生。

## 4. 匯入名單

老師後台的「名單管理」可貼上 Excel 或試算表資料。匯入後請確認學生姓名、學號、班級與 Gmail 是否正確。

## 5. 批閱與退件

老師在「老師後台」查看學生繳交狀態。退件時請填寫退件原因，系統會嘗試寄送 Email 給學生。
