# 🏠 套房管理系統 — 部署說明

> GitHub Pages + PWA + Google Apps Script 三層架構

---

## 架構總覽

```
使用者手機/電腦
    │  (PWA 安裝，離線可用靜態殼)
    ▼
GitHub Pages
    index.html  ← 前端介面（靜態）
    manifest.json
    sw.js       ← Service Worker
    icons/
    │
    │  (fetch POST，即時資料)
    ▼
Google Apps Script Web App
    程式碼.gs   ← 後端 API + doPost 路由
    │
    ▼
Google Sheets (您的資料)
```

---

## 📋 部署步驟

### Step 1：設定 Google Apps Script 後端

1. 開啟您的 Google Sheets（含 Rooms / Billings / Settlements 工作表）
2. 點選選單 **擴充功能 → Apps Script**
3. 刪除編輯器中所有內容，貼上本套件的 `程式碼.gs` 全文
4. 修改第 1 行的 `PHOTO_FOLDER_ID`（Google Drive 相片資料夾 ID）：
   ```javascript
   const PHOTO_FOLDER_ID = 'YOUR_FOLDER_ID_HERE';
   ```
5. 按 **儲存** (Ctrl+S)

#### 部署為 Web App
1. 點選右上角 **部署 → 新增部署**
2. 設定：
   - 類型：**網頁應用程式**
   - 說明：`套房管理 API v1`
   - 執行身分：**我（您的 Google 帳戶）**
   - 存取權：**所有人**（⚠️ 必須設為「所有人」，前端才能呼叫）
3. 按 **部署**，複製顯示的 **Web App URL**：
   ```
   https://script.google.com/macros/s/AKfycby.../exec
   ```
   > ⚠️ 每次修改程式碼後需重新部署（版本號會變）

---

### Step 2：部署前端到 GitHub Pages

1. **Fork 或建立新 Repository**（Public 或 Private 皆可）

2. **上傳本套件所有檔案**到 Repository 根目錄：
   ```
   index.html
   manifest.json
   sw.js
   icons/icon-192.png
   icons/icon-512.png
   README.md
   ```

3. 進入 Repository → **Settings → Pages**
4. Source 選 **Deploy from a branch**
5. Branch 選 **main**，目錄選 **/ (root)**
6. 按 **Save**，等待約 1 分鐘後取得網址：
   ```
   https://你的帳號.github.io/你的Repo名稱/
   ```

---

### Step 3：首次使用設定

1. 開啟 GitHub Pages 網址
2. 系統會彈出「連接 Google Sheets」對話框
3. 貼入 Step 1 取得的 Web App URL
4. 按「連接並儲存」

> **網址儲存於本機 localStorage，不會上傳至任何伺服器。**  
> 每台裝置需各自設定一次。

---

### Step 4：安裝為 PWA（選用）

#### Android / Chrome
1. 開啟 GitHub Pages 網址
2. Chrome 選單 → **「加入主畫面」**
3. 應用程式圖示會出現在桌面

#### iPhone / Safari
1. 開啟 GitHub Pages 網址（需使用 Safari）
2. 點選底部 **分享按鈕 (□↑)**
3. 選擇 **「加入主畫面」**
4. 點選「新增」

#### PC / Chrome
1. 網址列右側會出現 **安裝圖示 (⊕)**
2. 點選並確認安裝

---

## 🔐 安全性說明

| 面向 | 說明 |
|------|------|
| GAS Web App | 設為「所有人可存取」以讓前端呼叫，但資料在您的 Google 帳戶下 |
| API URL | 僅儲存在使用者本機 localStorage，不上傳 |
| 函數白名單 | `doPost` 僅允許呼叫 `ALLOWED` 列表中的函數 |
| 照片上傳 | 直接儲存至您指定的 Google Drive 資料夾 |
| GitHub Pages | 靜態檔案，無伺服器、無資料庫，GitHub 無法讀取您的資料 |

---

## 👥 多人使用說明

每位使用者需要：
1. 取得 GitHub Pages 網址（由您分享）
2. 首次開啟時輸入 Apps Script Web App URL（由您分享）
3. 各自在本機完成設定即可使用

**所有人共用同一個 Google Sheets 資料庫**，即時同步。

---

## 🔄 更新程式碼

### 更新前端
1. 修改 `index.html`
2. Push 到 GitHub
3. GitHub Pages 自動更新（約 1 分鐘）
4. 使用者重新整理頁面即可（PWA 會自動更新 Service Worker）

### 更新後端
1. 修改 `程式碼.gs`
2. 在 Apps Script 點選 **部署 → 管理部署 → 編輯（✏️）**
3. 版本選「**建立新版本**」
4. 按「**部署**」
5. > ⚠️ **Web App URL 不會改變**，不需要重新設定前端

---

## 🗂️ 工作表結構

系統需要以下 Google Sheets 工作表（首次執行 `initSettingsSheet` 會自動建立 Settings）：

| 工作表 | 用途 |
|--------|------|
| `Rooms` | 房間與租客基本資料 |
| `Billings` | 所有帳單記錄 |
| `Settlements` | 押金與台電結算 |
| `Taipower` | 台電帳單記錄 |
| `Settings` | 系統參數（自動建立） |

---

## ❓ 常見問題

**Q：出現「連線失敗」或「HTTP 401」？**  
→ GAS 部署時「存取權」需設為「所有人」，並確認已重新部署最新版本。

**Q：修改後端後功能失效？**  
→ Apps Script 修改後需重新部署新版本，URL 才會生效。

**Q：換手機後需要重新設定嗎？**  
→ 是的，每台裝置需各自在 localStorage 儲存 Web App URL（安全設計）。

**Q：可以同時讓多人使用嗎？**  
→ 可以，所有人共用同一 Google Sheets，即時同步。注意 GAS 有每日執行配額（免費版每日 6 分鐘），一般租管用量不會超過。

**Q：PWA 離線時可以用嗎？**  
→ 靜態介面可以離線顯示，但所有資料操作（抄表、開帳單等）需要網路連線才能與 Google Sheets 同步。

---

## 📞 重設連線

若需要更換 Apps Script URL（例如重新部署到新帳戶），進入：  
**系統設定頁 → 最下方「重設 API 連線」按鈕**

---

*套房管理系統 — 使用 Google Apps Script + GitHub Pages + PWA 架構*
