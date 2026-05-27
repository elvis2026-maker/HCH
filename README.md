# 台北市鍋爐壓力容器協會 官方網站

## 📁 目錄結構

```
tbpva/
├── index.html              # 首頁（主要入口）
├── css/
│   └── style.css           # 全站樣式（含RWD響應式）
├── js/
│   └── main.js             # 全站JavaScript（含LINE報名整合）
├── pages/
│   ├── about.html          # 協會簡介
│   ├── news.html           # 最新消息列表
│   ├── news-detail.html    # 消息內頁（待建立）
│   ├── download.html       # 下載專區
│   ├── member.html         # 會員服務（待建立）
│   ├── course-boiler.html  # 甲乙種鍋爐課程（待建立）
│   ├── course-pressure.html# 壓力容器課程（待建立）
│   ├── course-gas.html     # 高壓氣體課程（待建立）
│   ├── course-refresh.html # 鍋爐回訓課程（待建立）
│   ├── course-safety.html  # 丙種職安課程（待建立）
│   ├── history.html        # 發展沿革（待建立）
│   ├── mission.html        # 使命願景（待建立）
│   ├── org.html            # 組織架構（待建立）
│   ├── gallery.html        # 活動剪影（待建立）
│   ├── faq.html            # 常見問題（待建立）
│   ├── links.html          # 相關連結（待建立）
│   ├── privacy.html        # 隱私權政策（待建立）
│   └── terms.html          # 使用條款（待建立）
└── images/                 # 圖片目錄（放置協會相關圖片）
```

---

## 🚀 部署方式

### 方案一：GitHub Pages（免費，推薦）
1. 在 GitHub 建立 repository（例如：`tbpva-website`）
2. 將全部檔案上傳至 repository
3. 至 Settings → Pages → Source 選擇 `main` branch
4. 網址：`https://your-username.github.io/tbpva-website`
5. 可設定自訂網域（如 tbpva.org.tw）

### 方案二：自有主機/VPS（Apache/Nginx）
1. 將所有檔案上傳至主機 `/var/www/html/` 目錄
2. 設定 Apache/Nginx 指向該目錄
3. 設定 DNS 指向主機 IP

### 方案三：雲端靜態網站（Netlify/Vercel - 免費）
1. 至 [netlify.com](https://netlify.com) 或 [vercel.com](https://vercel.com) 免費申請
2. 連接 GitHub repository 或直接拖曳上傳資料夾
3. 自動部署，可設定自訂網域

---

## ⚙️ LINE 官方帳號設定

### 目前整合方式
網站使用 LINE URL Scheme 轉址整合，**不需要** LINE API 費用即可運作：
- 填寫報名表 → 自動產生含報名資訊的 LINE 訊息 → 開啟 LINE 完成傳送

### 步驟
1. 至 [LINE Official Account Manager](https://manager.line.biz/) 建立官方帳號
2. 取得帳號 ID（格式如 `@tbpva`）
3. 修改 `js/main.js` 中的 LINE URL：
   ```javascript
   // 找到以下兩行並更換為正確帳號 ID
   const lineUrl = `https://line.me/R/oaMessage/@tbpva/?${encoded}`;
   window.open('https://line.me/R/ti/p/@tbpva', '_blank');
   ```
4. 修改 HTML 中的 LINE 連結（搜尋 `@tbpva` 全部取代）

### 進階整合（需付費帳號）
如需自動回覆、報名資料庫記錄，可串接 LINE Messaging API：
- 申請 LINE Developers 帳號
- 建立 Messaging API Channel
- 使用 Google Sheets / Airtable 作為輕量資料庫

---

## ✏️ 內容更新方式

### 新增消息
在 `pages/news.html` 複製一個 `.news-list-item` 區塊，修改日期、標題、摘要。

### 修改課程資訊
直接編輯 `index.html` 中各 `.course-card` 的內容。

### 修改聯絡資訊
搜尋 `02-2552-0255` 全域取代為正確電話；搜尋 `tbpva125@gmail.com` 取代為正確信箱。

### 新增下載文件
在 `pages/download.html` 的 `.dl-grid` 中新增 `.dl-item`，並將 `href="#"` 改為實際文件路徑。

---

## 🎨 設計規格

| 項目 | 說明 |
|------|------|
| 主色 | `#1a3a5c`（深藍色） |
| 輔色 | `#c8922a`（金色） |
| LINE | `#06C755`（LINE 綠） |
| 字型 | Noto Serif TC（標題） + Noto Sans TC（內文） |
| 斷點 | 480px / 768px / 1024px |
| 框架 | 純 HTML + CSS + JS（無需 Node.js 或任何框架） |

---

## 📋 待辦事項

- [ ] 確認並更新 LINE 官方帳號 ID
- [ ] 更換協會實際照片至 `images/` 目錄
- [ ] 建立各課程詳細頁面
- [ ] 填入實際下載文件 PDF 連結
- [ ] 設定 Google Maps 正確嵌入連結
- [ ] 建立 `sitemap.xml` 提交 Google Search Console
- [ ] 設定 SSL 憑證（HTTPS）
- [ ] 設定 DNS A Record 指向主機

---

## 💡 技術說明

- **純靜態網站**：無需資料庫、無需伺服器端語言，任何主機皆可部署
- **RWD 響應式**：支援手機（320px）至桌機（1440px+）
- **SEO 友善**：包含 meta 標籤、語意化 HTML5
- **無障礙**：ARIA 標籤、鍵盤操作支援
- **效能**：Google Fonts 非同步載入，純 CSS 動畫

---

聯絡技術支援：請洽負責建置之網頁工程師
