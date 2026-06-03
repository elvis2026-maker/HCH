# HCH 鴻琪股份有限公司 — 新版企業網站
## 部署說明 Deployment Guide

---

### 系統需求 System Requirements

- **Web Server**: IIS 6.0 以上 (Windows Server 2003+)
- **ASP**: Classic ASP (VBScript) — 需啟用 ASP 功能
- **Database**: Microsoft Access 2000+ (.mdb)
- **ODBC Driver**: Microsoft Access Database Engine (已安裝 Office 或單獨安裝)
- **Fonts**: Google Fonts (需連線), 或可離線替換為系統字體

---

### 目錄結構 Directory Structure

```
/
├── index.asp           ← 首頁 Homepage
├── products.asp        ← 產品列表頁 Products listing
├── search.asp          ← 料號搜尋頁 Part number search
├── about.asp           ← 公司介紹頁 About page
├── contact.asp         ← 聯絡我們頁 Contact page
├── catalog.asp         ← 型錄下載頁 Catalog download
├── message_save.asp    ← AJAX 留言儲存 Message save endpoint
├── 404.asp             ← 404 錯誤頁
├── db_setup.asp        ← 資料庫初始化 (執行一次後刪除)
├── web.config          ← IIS 設定檔
├── inc_header.asp      ← 共用頁首 (include)
├── inc_footer.asp      ← 共用頁尾+客服 widget (include)
│
├── css/
│   └── style.css       ← 主要樣式表
│
├── images/             ← 圖片資源 (從舊站複製)
│   ├── hchfavicon.ico
│   ├── hchicon.png
│   ├── HCHLINE.jpg
│   └── index2023.jpg
│
├── database/
│   └── data_new.mdb    ← Access 資料庫 (含產品資料 + 留言表)
│
├── html/
│   ├── CON_01.pdf      ← 綜合型錄 PDF
│   ├── CON_02.pdf
│   ├── CON_03.pdf
│   ├── CON_04.pdf
│   └── single/         ← 單品規格圖檔 PDF (2,900+ 筆)
│       ├── STS-05.pdf
│       └── ...
│
├── data/
│   └── count.txt       ← 訪客計數器
│
└── admin/
    └── messages.asp    ← 留言管理後台
```

---

### 部署步驟 Setup Steps

1. **上傳所有檔案** 至您的 IIS 網站根目錄

2. **設定 IIS**:
   - 啟用 ASP 功能 (Enable ASP in IIS Features)
   - 設定預設文件為 `index.asp`
   - 確認 Application Pool 使用 Classic ASP 或 Integrated mode

3. **資料庫初始設定**:
   - 開啟瀏覽器訪問: `http://yoursite/db_setup.asp?key=hchsetup2024`
   - 此腳本會建立 `messages` 資料表（用於儲存留言）
   - 執行成功後，**請刪除或重新命名 db_setup.asp**

4. **資料夾權限**:
   - `database/` 資料夾需要 **讀寫權限** (IIS User / IUSR 帳號)
   - `data/` 資料夾需要 **讀寫權限** (訪客計數器)
   - `html/` 資料夾需要 **讀取權限**

5. **PDF 規格圖檔**:
   - 將所有單品 PDF 放入 `html/single/` 資料夾
   - 檔名格式與資料庫 `search` 表的 `LINK` 欄位對應
   - 例如: LINK = "STS-05.pdf" → 檔案放在 `html/single/STS-05.pdf`

---

### 功能說明 Features

| 功能 | 說明 |
|------|------|
| 響應式設計 | 手機 / 平板 / 電腦均可正常顯示 |
| 料號搜尋 | 搜尋資料庫中 2,900+ 筆料號，顯示 PDF 連結 |
| 產品分類 | 13 大類產品，對應 Access 資料庫各資料表 |
| PDF 預覽 | 直接在頁面內嵌顯示 PDF，亦可另開下載 |
| 線上留言 | 表單送出後儲存至 Access 資料庫 messages 表 |
| 管理後台 | /admin/messages.asp — 查看/刪除留言 |
| LINE 整合 | 頁面多處整合 LINE 官方帳號 @307puwgm |
| 客服 Widget | 右下角浮動客服按鈕，包含 LINE/Email/電話/留言 |
| 訪客計數 | 首頁顯示網站訪客計數 (data/count.txt) |
| 型錄下載 | 提供 4 本綜合型錄 PDF 下載 |
| 新聞跑馬燈 | 首頁頂部顯示新產品捲動訊息 |

---

### 管理後台 Admin Panel

- **URL**: `/admin/messages.asp`
- **帳號**: `hchlogin`
- **密碼**: `looklogin`
- 功能: 查看客戶留言、刪除留言

> ⚠️ 建議上線後修改管理員帳號密碼 (在 admin/messages.asp 第 11-12 行)

---

### LINE 官方帳號設定

目前設定的 LINE 帳號: **@307puwgm**
- 首頁 CTA 按鈕
- 頁尾聯絡資訊
- 客服 Widget
- 關於我們頁面

如需更換，全文搜尋 `307puwgm` 並替換。

---

### 新增 PDF 規格圖檔

1. 將 PDF 檔案放入 `html/single/` 資料夾
2. 更新 Access 資料庫的 `search` 表，在 `LINK` 欄位填入 PDF 檔名
3. 確認 `LINK` 欄位值與實際檔案名稱完全相符（區分大小寫）

---

### 技術規格 Technical Specs

- Language: Classic ASP / VBScript
- Database: Microsoft Access 2003 (.mdb)
- Charset: UTF-8
- CSS: Custom (無框架依賴)
- Icons: Font Awesome 6.4
- Fonts: Google Fonts (Rajdhani + Noto Sans TC)
- Browser: IE 11+, Chrome, Firefox, Safari, Edge

---

Copyright © 1984-2026 HCH 鴻琪股份有限公司. All rights reserved.
Website: https://www.hchtwn.com.tw
