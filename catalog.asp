<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
PageTitle = "型錄下載 Catalog Download"
%>
<%@include file="inc_header.asp"%>

<!-- BREADCRUMB -->
<div class="breadcrumb">
  <div class="container">
    <ul class="breadcrumb-list">
      <li><a href="index.asp">首頁</a></li>
      <li>型錄下載</li>
    </ul>
  </div>
</div>

<!-- PAGE HERO -->
<div class="page-hero">
  <div class="container">
    <h1><i class="fas fa-download" style="font-size:32px;margin-right:14px;opacity:0.8;"></i>型錄下載 Catalog Download</h1>
    <p>下載 HCH 最新產品型錄 PDF，了解完整產品規格與系列資訊。</p>
  </div>
</div>

<section class="section">
  <div class="container">

    <div class="section-header">
      <div class="section-tag">DOWNLOAD</div>
      <h2 class="section-title">產品型錄</h2>
      <p class="section-desc">點擊下方型錄封面可直接開啟 PDF 檔案，或按下載按鈕儲存至您的裝置。需要 Adobe Reader 開啟 PDF 文件。</p>
      <div class="section-divider"></div>
    </div>

    <!-- CATALOG CARDS -->
    <div class="catalog-grid" style="grid-template-columns:repeat(auto-fill,minmax(240px,1fr));gap:24px;margin-bottom:50px;">
      <%
      Dim catalogs(3,2)
      catalogs(0,0) = "綜合型錄 Vol. 01"
      catalogs(0,1) = "html/CON_01.pdf"
      catalogs(0,2) = "包含 MICRO CONNECTORS、MINI DIN CONNECTORS、TERMINAL CONNECTORS 等系列"

      catalogs(1,0) = "綜合型錄 Vol. 02"
      catalogs(1,1) = "html/CON_02.pdf"
      catalogs(1,2) = "包含 FIX TERMINALS、FUSE HOLDER、BATTERY CONNECTORS、LENSES 等系列"

      catalogs(2,0) = "綜合型錄 Vol. 03"
      catalogs(2,1) = "html/CON_03.pdf"
      catalogs(2,2) = "包含 PHOTOMOS RELAYS、U.S. RECEPTACLE 等系列"

      catalogs(3,0) = "綜合型錄 Vol. 04"
      catalogs(3,1) = "html/CON_04.pdf"
      catalogs(3,2) = "包含最新產品系列及更新規格資訊"

      Dim ci
      For ci = 0 To 3
        Response.Write "<div style=""background:#fff;border:1px solid #E0E0E0;border-radius:10px;overflow:hidden;box-shadow:0 2px 12px rgba(0,0,0,0.07);transition:box-shadow 0.2s,transform 0.2s;"" onmouseover=""this.style.boxShadow='0 8px 30px rgba(0,0,0,0.13)';this.style.transform='translateY(-4px)'"" onmouseout=""this.style.boxShadow='0 2px 12px rgba(0,0,0,0.07)';this.style.transform=''"">"
        Response.Write "<a href=""" & catalogs(ci,1) & """ target=""_blank"" style=""display:block;height:180px;background:linear-gradient(135deg,#0D0D0D,#1a0000,#CC0000);display:flex;align-items:center;justify-content:center;flex-direction:column;gap:10px;"">"
        Response.Write "<i class=""fas fa-file-pdf"" style=""font-size:52px;color:#fff;opacity:0.9;""></i>"
        Response.Write "<span style=""font-family:'Rajdhani',sans-serif;font-size:20px;font-weight:700;color:#fff;letter-spacing:0.05em;"">HCH CATALOG</span>"
        Response.Write "<span style=""font-size:11px;color:rgba(255,255,255,0.6);letter-spacing:0.1em;"">VOL. 0" & (ci+1) & "</span>"
        Response.Write "</a>"
        Response.Write "<div style=""padding:18px 20px;"">"
        Response.Write "<h4 style=""font-size:15px;font-weight:700;margin-bottom:6px;color:#111;"">" & catalogs(ci,0) & "</h4>"
        Response.Write "<p style=""font-size:12px;color:#777;line-height:1.6;margin-bottom:14px;"">" & catalogs(ci,2) & "</p>"
        Response.Write "<div style=""display:flex;gap:10px;"">"
        Response.Write "<a href=""" & catalogs(ci,1) & """ target=""_blank"" style=""flex:1;display:inline-flex;align-items:center;justify-content:center;gap:6px;background:#CC0000;color:#fff;padding:9px 0;border-radius:4px;font-size:13px;font-weight:700;transition:background 0.2s;"" onmouseover=""this.style.background='#990000'"" onmouseout=""this.style.background='#CC0000'"">"
        Response.Write "<i class=""fas fa-eye""></i> 開啟預覽</a>"
        Response.Write "<a href=""" & catalogs(ci,1) & """ download style=""flex:1;display:inline-flex;align-items:center;justify-content:center;gap:6px;background:#003399;color:#fff;padding:9px 0;border-radius:4px;font-size:13px;font-weight:700;transition:background 0.2s;"" onmouseover=""this.style.background='#002277'"" onmouseout=""this.style.background='#003399'"">"
        Response.Write "<i class=""fas fa-download""></i> 下載</a>"
        Response.Write "</div></div></div>"
      Next
      %>
    </div>

    <!-- SINGLE PRODUCT PDF SECTION -->
    <div style="background:#F5F5F5;border-radius:10px;padding:36px;margin-bottom:40px;">
      <div style="display:flex;align-items:flex-start;gap:20px;flex-wrap:wrap;">
        <div style="flex:1;min-width:260px;">
          <h3 style="font-size:20px;font-weight:700;margin-bottom:10px;color:#111;">
            <i class="fas fa-file-alt" style="color:#CC0000;margin-right:10px;"></i>單品規格圖檔
          </h3>
          <p style="font-size:14px;color:#555;line-height:1.8;margin-bottom:16px;">
            每個料號均有對應的 PDF 規格圖檔。請使用料號搜尋功能找到您需要的產品規格，
            或瀏覽產品分類頁面直接下載單品 PDF。
          </p>
          <div style="display:flex;gap:12px;flex-wrap:wrap;">
            <a href="search.asp" class="btn-primary" style="display:inline-flex;align-items:center;gap:8px;">
              <i class="fas fa-search"></i> 搜尋料號 PDF
            </a>
            <a href="products.asp?cat=smd" style="display:inline-flex;align-items:center;gap:8px;background:#003399;color:#fff;padding:12px 22px;border-radius:4px;font-weight:700;font-size:14px;transition:background 0.2s;" onmouseover="this.style.background='#002277'" onmouseout="this.style.background='#003399'">
              <i class="fas fa-th-list"></i> 瀏覽產品目錄
            </a>
          </div>
        </div>
        <div style="background:#fff;border:1px solid #E0E0E0;border-radius:8px;padding:20px;flex-shrink:0;min-width:220px;">
          <h5 style="font-size:13px;font-weight:700;color:#333;margin-bottom:12px;"><i class="fas fa-info-circle" style="color:#003399;margin-right:6px;"></i>PDF 開啟說明</h5>
          <ul style="font-size:12px;color:#666;line-height:2;">
            <li>• 需安裝 Adobe Reader 或其他 PDF 閱讀器</li>
            <li>• 手機可直接使用內建 PDF 閱讀器開啟</li>
            <li>• 如無法開啟請聯絡我們取得協助</li>
          </ul>
          <a href="http://get.adobe.com/tw/reader/" target="_blank" rel="noopener" style="display:inline-flex;align-items:center;gap:6px;font-size:12px;color:#CC0000;font-weight:600;margin-top:10px;">
            <i class="fas fa-download"></i> 下載 Adobe Reader
          </a>
        </div>
      </div>
    </div>

    <!-- NEED HELP -->
    <div style="background:linear-gradient(135deg,#111,#1a0000);border-radius:10px;padding:36px;color:#fff;text-align:center;">
      <h3 style="font-family:'Rajdhani',sans-serif;font-size:28px;font-weight:700;margin-bottom:10px;">找不到您需要的型錄？</h3>
      <p style="font-size:14px;color:rgba(255,255,255,0.7);margin-bottom:24px;">歡迎透過以下方式聯繫我們，業務人員將為您提供最完整的產品資料。</p>
      <div style="display:flex;gap:14px;justify-content:center;flex-wrap:wrap;">
        <a href="contact.asp" class="btn-primary" style="display:inline-flex;align-items:center;gap:8px;padding:13px 26px;font-size:14px;">
          <i class="fas fa-paper-plane"></i> 線上留言
        </a>
        <a href="https://line.me/R/ti/p/@307puwgm" target="_blank" rel="noopener" class="btn-line" style="padding:13px 26px;font-size:14px;border-radius:4px;">
          <svg width="18" height="18" viewBox="0 0 24 24" fill="currentColor"><path d="M19.952 12.375C19.952 8.33 15.9 5.03 11 5.03c-4.9 0-8.952 3.3-8.952 7.345 0 3.63 3.22 6.674 7.572 7.25.295.064.696.196.797.45.092.23.06.59.03.822 0 0-.107.639-.13.775-.04.23-.185.9.79.49.974-.408 5.254-3.094 7.17-5.297 1.32-1.452 1.673-2.904 1.673-4.49z"/></svg>
          LINE 洽詢
        </a>
        <a href="mailto:hch@hchtwn.com.tw" style="display:inline-flex;align-items:center;gap:8px;background:rgba(255,255,255,0.1);color:#fff;padding:13px 26px;border-radius:4px;font-size:14px;font-weight:700;border:1px solid rgba(255,255,255,0.2);transition:background 0.2s;" onmouseover="this.style.background='rgba(255,255,255,0.2)'" onmouseout="this.style.background='rgba(255,255,255,0.1)'">
          <i class="fas fa-envelope"></i> Email
        </a>
      </div>
    </div>

  </div>
</section>

<%@include file="inc_footer.asp"%>
