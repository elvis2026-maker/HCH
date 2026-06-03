<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
PageTitle = "公司介紹 About Us"
%>
<%@include file="inc_header.asp"%>

<!-- BREADCRUMB -->
<div class="breadcrumb">
  <div class="container">
    <ul class="breadcrumb-list">
      <li><a href="index.asp">首頁</a></li>
      <li>公司介紹</li>
    </ul>
  </div>
</div>

<!-- PAGE HERO -->
<div class="page-hero">
  <div class="container">
    <h1><i class="fas fa-building" style="font-size:32px;margin-right:14px;opacity:0.8;"></i>公司介紹 About HCH</h1>
    <p>用心創新第一，品質服務優先。HCH — 值得您信賴的電子零件夥伴。</p>
  </div>
</div>

<!-- ABOUT MAIN -->
<section class="section">
  <div class="container">
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:50px;align-items:start;">
      <div>
        <div class="section-tag">ABOUT US</div>
        <h2 class="section-title" style="text-align:left;margin-bottom:18px;">鴻琪股份有限公司</h2>
        <div class="section-divider" style="margin:0 0 24px 0;"></div>
        <p style="font-size:15px;line-height:1.9;color:#444;margin-bottom:16px;">
          鴻琪股份有限公司（HCH）創立於 <strong>1984 年</strong>，是一家專業從事電子零件製造與銷售的企業，
          主要產品包含各類開關、連接器、感應器及相關電子零件。
        </p>
        <p style="font-size:15px;line-height:1.9;color:#444;margin-bottom:16px;">
          HCH was established in 1984, specializing in the manufacturing and distribution of electronic components,
          including switches, connectors, sensors and related electronic parts.
        </p>
        <p style="font-size:15px;line-height:1.9;color:#444;margin-bottom:24px;">
          40 年以上的製造經驗，通過 ISO 及 SGS 國際認證，產品品質達到國際標準，
          並外銷至多個國家地區，建立廣泛的全球客戶網絡。
        </p>
        <div style="display:flex;gap:20px;flex-wrap:wrap;margin-bottom:28px;">
          <div style="text-align:center;background:#FFF5F5;border:2px solid #CC0000;border-radius:8px;padding:20px 24px;min-width:100px;">
            <div style="font-family:'Rajdhani',sans-serif;font-size:36px;font-weight:700;color:#CC0000;line-height:1;">40+</div>
            <div style="font-size:12px;color:#666;margin-top:4px;">年製造經驗<br>Years Experience</div>
          </div>
          <div style="text-align:center;background:#F0F4FF;border:2px solid #003399;border-radius:8px;padding:20px 24px;min-width:100px;">
            <div style="font-family:'Rajdhani',sans-serif;font-size:36px;font-weight:700;color:#003399;line-height:1;">2,900+</div>
            <div style="font-size:12px;color:#666;margin-top:4px;">料號品項<br>Part Numbers</div>
          </div>
          <div style="text-align:center;background:#F0FFF4;border:2px solid #2E7D32;border-radius:8px;padding:20px 24px;min-width:100px;">
            <div style="font-family:'Rajdhani',sans-serif;font-size:36px;font-weight:700;color:#2E7D32;line-height:1;">ISO</div>
            <div style="font-size:12px;color:#666;margin-top:4px;">國際認證<br>Certified</div>
          </div>
        </div>
        <a href="contact.asp" class="btn-primary" style="display:inline-flex;align-items:center;gap:8px;">
          <i class="fas fa-envelope"></i> 立即聯絡我們
        </a>
        <a href="search.asp" style="display:inline-flex;align-items:center;gap:8px;margin-left:14px;color:#003399;font-weight:600;font-size:14px;">
          <i class="fas fa-search"></i> 搜尋料號
        </a>
      </div>

      <div>
        <div style="background:#F5F5F5;border-radius:10px;overflow:hidden;margin-bottom:20px;">
          <% If False Then %>
          <img src="images/company2022.gif" alt="HCH 公司" style="width:100%;">
          <% Else %>
          <div style="height:220px;display:flex;align-items:center;justify-content:center;background:linear-gradient(135deg,#111,#1a0000);color:rgba(255,255,255,0.6);flex-direction:column;gap:10px;">
            <div style="font-family:'Rajdhani',sans-serif;font-size:64px;font-weight:700;color:#CC0000;">HCH</div>
            <div style="font-size:14px;letter-spacing:0.1em;">鴻琪股份有限公司</div>
            <div style="font-size:11px;color:rgba(255,255,255,0.4);">EST. 1984</div>
          </div>
          <% End If %>
        </div>

        <!-- CERTIFICATIONS -->
        <div style="background:#fff;border:1px solid #E0E0E0;border-radius:8px;padding:20px;margin-bottom:20px;">
          <h4 style="font-size:15px;font-weight:700;margin-bottom:14px;color:#111;"><i class="fas fa-certificate" style="color:#CC0000;margin-right:8px;"></i>認證 Certifications</h4>
          <div style="display:flex;gap:14px;flex-wrap:wrap;">
            <div style="flex:1;min-width:100px;text-align:center;background:#F5F5F5;border-radius:6px;padding:14px;">
              <div style="font-size:28px;margin-bottom:6px;">🏅</div>
              <div style="font-size:13px;font-weight:700;">ISO</div>
              <div style="font-size:11px;color:#666;">國際標準認證</div>
            </div>
            <div style="flex:1;min-width:100px;text-align:center;background:#F5F5F5;border-radius:6px;padding:14px;">
              <div style="font-size:28px;margin-bottom:6px;">✅</div>
              <div style="font-size:13px;font-weight:700;">SGS</div>
              <div style="font-size:11px;color:#666;">品質認證</div>
            </div>
            <div style="flex:1;min-width:100px;text-align:center;background:#F5F5F5;border-radius:6px;padding:14px;">
              <div style="font-size:28px;margin-bottom:6px;">🌍</div>
              <div style="font-size:13px;font-weight:700;">全球出口</div>
              <div style="font-size:11px;color:#666;">Global Export</div>
            </div>
          </div>
        </div>

        <!-- LINE QR CODE -->
        <div style="background:#E8FAF0;border:1.5px solid #06C755;border-radius:8px;padding:16px 20px;display:flex;align-items:center;gap:16px;">
          <div style="background:#06C755;width:52px;height:52px;border-radius:50%;display:flex;align-items:center;justify-content:center;flex-shrink:0;">
            <svg width="30" height="30" viewBox="0 0 24 24" fill="white"><path d="M19.952 12.375C19.952 8.33 15.9 5.03 11 5.03c-4.9 0-8.952 3.3-8.952 7.345 0 3.63 3.22 6.674 7.572 7.25.295.064.696.196.797.45.092.23.06.59.03.822 0 0-.107.639-.13.775-.04.23-.185.9.79.49.974-.408 5.254-3.094 7.17-5.297 1.32-1.452 1.673-2.904 1.673-4.49z"/></svg>
          </div>
          <div>
            <div style="font-weight:700;color:#111;margin-bottom:4px;">LINE 官方帳號</div>
            <div style="font-size:12px;color:#555;margin-bottom:8px;">@307puwgm — 加入好友即時洽詢</div>
            <a href="https://line.me/R/ti/p/@307puwgm" target="_blank" rel="noopener" style="display:inline-block;background:#06C755;color:#fff;padding:7px 18px;border-radius:4px;font-size:13px;font-weight:700;">加入 LINE</a>
          </div>
        </div>
      </div>
    </div>
  </div>
</section>

<!-- COMPANY HISTORY TIMELINE -->
<section class="section section-alt">
  <div class="container">
    <div class="section-header">
      <div class="section-tag">HISTORY</div>
      <h2 class="section-title">發展歷程</h2>
      <p class="section-desc">從 1984 年創立至今，持續成長與創新。</p>
      <div class="section-divider"></div>
    </div>
    <div style="max-width:700px;margin:0 auto;">
      <%
      Dim years, events
      years  = Array("1984","1990s","2000s","2010s","2020+")
      events = Array("鴻琪股份有限公司正式創立，開始生產各類電子開關零件。","業務擴展，產品線涵蓋更多類型的開關及連接器，開始外銷業務。","通過 ISO 及 SGS 國際認證，品質管理達到國際標準，全球出口規模擴大。","持續研發新產品，引進 SMD 系列、FPC 連接器等先進電子零件，強化產品競爭力。","成立 40 週年，持續推出新產品系列，擴充線上銷售與數位化服務。")
      Dim ti
      For ti = 0 To UBound(years)
        Dim isLeft : If ti Mod 2 = 0 Then isLeft = True Else isLeft = False
        Response.Write "<div style=""display:flex;gap:0;margin-bottom:0;align-items:stretch;"">"
        If isLeft Then
          Response.Write "<div style=""flex:1;text-align:right;padding:0 24px 28px 0;"">"
          Response.Write "<div style=""background:#fff;border:1px solid #E0E0E0;border-radius:8px;padding:16px 20px;display:inline-block;max-width:280px;box-shadow:0 2px 8px rgba(0,0,0,0.06);"">"
          Response.Write "<div style=""font-family:'Rajdhani',sans-serif;font-size:22px;font-weight:700;color:#CC0000;margin-bottom:6px;"">" & years(ti) & "</div>"
          Response.Write "<div style=""font-size:13px;color:#555;line-height:1.7;"">" & events(ti) & "</div>"
          Response.Write "</div></div>"
          Response.Write "<div style=""display:flex;flex-direction:column;align-items:center;"">"
          Response.Write "<div style=""width:18px;height:18px;background:#CC0000;border-radius:50%;border:3px solid #fff;box-shadow:0 0 0 3px #CC0000;flex-shrink:0;margin-top:14px;""></div>"
          If ti < UBound(years) Then Response.Write "<div style=""flex:1;width:2px;background:#E0E0E0;""></div>"
          Response.Write "</div>"
          Response.Write "<div style=""flex:1;""></div>"
        Else
          Response.Write "<div style=""flex:1;""></div>"
          Response.Write "<div style=""display:flex;flex-direction:column;align-items:center;"">"
          Response.Write "<div style=""width:18px;height:18px;background:#003399;border-radius:50%;border:3px solid #fff;box-shadow:0 0 0 3px #003399;flex-shrink:0;margin-top:14px;""></div>"
          If ti < UBound(years) Then Response.Write "<div style=""flex:1;width:2px;background:#E0E0E0;""></div>"
          Response.Write "</div>"
          Response.Write "<div style=""flex:1;padding:0 0 28px 24px;"">"
          Response.Write "<div style=""background:#fff;border:1px solid #E0E0E0;border-radius:8px;padding:16px 20px;display:inline-block;max-width:280px;box-shadow:0 2px 8px rgba(0,0,0,0.06);"">"
          Response.Write "<div style=""font-family:'Rajdhani',sans-serif;font-size:22px;font-weight:700;color:#003399;margin-bottom:6px;"">" & years(ti) & "</div>"
          Response.Write "<div style=""font-size:13px;color:#555;line-height:1.7;"">" & events(ti) & "</div>"
          Response.Write "</div></div>"
        End If
        Response.Write "</div>"
      Next
      %>
    </div>
  </div>
</section>

<!-- CORE VALUES -->
<section class="section">
  <div class="container">
    <div class="section-header">
      <div class="section-tag">VALUES</div>
      <h2 class="section-title">核心理念</h2>
      <div class="section-divider"></div>
    </div>
    <div class="features">
      <div class="feature-card">
        <div class="feature-icon"><i class="fas fa-lightbulb"></i></div>
        <h3>用心創新第一</h3>
        <p>持續研發創新，引進最新技術，提供符合市場需求的高品質電子零件。</p>
      </div>
      <div class="feature-card">
        <div class="feature-icon"><i class="fas fa-star"></i></div>
        <h3>品質服務優先</h3>
        <p>嚴格的品質管理制度，通過 ISO / SGS 認證，確保每批產品符合標準。</p>
      </div>
      <div class="feature-card">
        <div class="feature-icon"><i class="fas fa-leaf"></i></div>
        <h3>提昇工作品質</h3>
        <p>不斷改善生產流程，提升工作環境品質，打造專業高效的製造團隊。</p>
      </div>
      <div class="feature-card">
        <div class="feature-icon"><i class="fas fa-recycle"></i></div>
        <h3>做好生活環保</h3>
        <p>重視環境保護，推行綠色製造，生產符合 RoHS 等環保規範的產品。</p>
      </div>
    </div>
  </div>
</section>

<%@include file="inc_footer.asp"%>
