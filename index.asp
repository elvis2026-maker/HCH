<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
' ===== VISIT COUNTER =====
Dim CountFile, FileObject, Out, counter
CountFile = Server.MapPath("data/count.txt")
On Error Resume Next
Set FileObject = Server.CreateObject("Scripting.FileSystemObject")
Set Out = FileObject.OpenTextFile(CountFile, 1, FALSE, FALSE)
If Err.Number = 0 Then counter = CLng(Out.ReadLine) Else counter = 0
Out.Close
Application.Lock
counter = counter + 1
Set Out = FileObject.CreateTextFile(CountFile, TRUE, FALSE)
Out.WriteLine(counter)
Out.Close
Application.Unlock
Set FileObject = Nothing
On Error GoTo 0

' ===== DB CONNECTION =====
Function GetConn()
  Dim conn, strCnn
  Set conn = Server.CreateObject("ADODB.Connection")
  strCnn = "Driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath("database/data_new.mdb")
  conn.Open strCnn
  Set GetConn = conn
End Function

PageTitle = "首頁"
%>
<%@include file="inc_header.asp"%>

<!-- ===== NEWS TICKER ===== -->
<div class="news-bar">
  <div class="container">
    <div class="news-bar-inner">
      <span class="news-label">NEW</span>
      <div class="news-ticker-wrap">
        <div class="news-ticker">
          <%
          Dim connN, rsN
          Set connN = GetConn()
          Set rsN = connN.Execute("SELECT * FROM news WHERE name IS NOT NULL AND name <> '' ORDER BY id DESC")
          Dim tickerHTML : tickerHTML = ""
          Do While Not rsN.EOF
            tickerHTML = tickerHTML & "<a href=""products.asp?cat=news&link=" & rsN("link") & """>" & rsN("name") & "</a>"
            rsN.MoveNext
          Loop
          rsN.Close: Set rsN = Nothing
          connN.Close: Set connN = Nothing
          Response.Write tickerHTML & tickerHTML ' double for seamless loop
          %>
        </div>
      </div>
    </div>
  </div>
</div>

<!-- ===== HERO SECTION ===== -->
<section class="hero">
  <div class="container">
    <div class="hero-inner">
      <div class="hero-content">
        <div class="hero-badge">Since 1984 · ISO Certified · SGS Certified</div>
        <h1><span>HCH</span> 鴻琪<br>專業電子零件</h1>
        <p class="hero-desc">
          鴻琪股份有限公司創立於 1984 年，專業製造各類電子開關、連接器及感應器，
          通過 ISO 及 SGS 國際認證，產品外銷全球，品質服務優先。
        </p>
        <div class="hero-ctas">
          <a href="products.asp?cat=smd" class="btn-primary">瀏覽產品目錄</a>
          <a href="search.asp" class="btn-outline"><i class="fas fa-search" style="margin-right:6px;"></i>料號搜尋</a>
          <a href="https://line.me/R/ti/p/@307puwgm" target="_blank" rel="noopener" class="btn-outline" style="background:rgba(6,199,85,0.15);border-color:#06C755;color:#fff;">
            <svg style="width:16px;height:16px;fill:currentColor;margin-right:6px;" viewBox="0 0 24 24"><path d="M19.952 12.375C19.952 8.33 15.9 5.03 11 5.03c-4.9 0-8.952 3.3-8.952 7.345 0 3.63 3.22 6.674 7.572 7.25.295.064.696.196.797.45.092.23.06.59.03.822 0 0-.107.639-.13.775-.04.23-.185.9.79.49.974-.408 5.254-3.094 7.17-5.297 1.32-1.452 1.673-2.904 1.673-4.49z"/></svg>
            LINE 洽詢
          </a>
        </div>
        <div class="hero-stats">
          <div class="hero-stat">
            <strong>40+</strong>
            <span>Years Experience</span>
          </div>
          <div class="hero-stat">
            <strong>2,900+</strong>
            <span>Part Numbers</span>
          </div>
          <div class="hero-stat">
            <strong>ISO</strong>
            <span>Certified Quality</span>
          </div>
          <div class="hero-stat">
            <strong><%=counter%></strong>
            <span>Website Visitors</span>
          </div>
        </div>
      </div>
      <div class="hero-visual">
        <div class="hero-img-wrapper">
          <% If FileExists("images/index2023.jpg") Then %>
          <img src="images/index2023.jpg" alt="HCH 鴻琪 電子零件" style="max-height:220px;margin:0 auto;">
          <% Else %>
          <div class="hero-placeholder">
            <div class="icon">⚡</div>
            <div>HCH 電子零件</div>
            <div style="font-size:11px;opacity:0.6;">Electronic Components</div>
          </div>
          <% End If %>
        </div>
        <% If FileExists("images/40th.jpg") Then %>
        <div style="margin-top:14px;border-radius:6px;overflow:hidden;border:1px solid rgba(255,255,255,0.1);">
          <img src="images/40th.jpg" alt="HCH 40週年" style="width:100%;max-height:100px;object-fit:cover;">
        </div>
        <% End If %>
      </div>
    </div>
  </div>
</section>

<!-- ===== FEATURES ===== -->
<section class="section section-alt" style="padding:50px 0;">
  <div class="container">
    <div class="features">
      <div class="feature-card">
        <div class="feature-icon"><i class="fas fa-certificate"></i></div>
        <h3>ISO / SGS 認證</h3>
        <p>通過國際認證，品質管理達到國際標準，每批產品均經嚴格檢驗。</p>
      </div>
      <div class="feature-card">
        <div class="feature-icon"><i class="fas fa-industry"></i></div>
        <h3>專業製造</h3>
        <p>40 年以上製造經驗，擁有完整設計研發與生產能力，快速回應客戶需求。</p>
      </div>
      <div class="feature-card">
        <div class="feature-icon"><i class="fas fa-globe"></i></div>
        <h3>全球出口</h3>
        <p>產品外銷多個國家地區，建立廣泛的全球銷售網絡，服務各地客戶。</p>
      </div>
      <div class="feature-card">
        <div class="feature-icon"><i class="fas fa-cogs"></i></div>
        <h3>多元產品線</h3>
        <p>涵蓋開關、連接器、感應器等多種電子零件，一站式採購解決方案。</p>
      </div>
    </div>
  </div>
</section>

<!-- ===== PRODUCTS SECTION ===== -->
<section class="section">
  <div class="container">
    <div class="section-header">
      <div class="section-tag">PRODUCTS</div>
      <h2 class="section-title">產品目錄</h2>
      <p class="section-desc">鴻琪提供多種高品質電子零件，包含開關、連接器、感應器等，滿足各種應用需求。</p>
      <div class="section-divider"></div>
    </div>
    <div class="product-categories">

      <!-- SMD -->
      <div class="cat-card">
        <div class="cat-card-header">
          <span class="cat-icon">🔲</span>
          <span class="cat-title">SMD 觸控開關</span>
        </div>
        <div class="cat-card-body">
          <%
          Dim cSmd, rSmd
          Set cSmd = GetConn()
          Set rSmd = cSmd.Execute("SELECT TOP 6 * FROM smd WHERE name IS NOT NULL AND name <> '' ORDER BY id")
          Do While Not rSmd.EOF
            Dim nm : nm = rSmd("name")
            Dim lk : lk = rSmd("link")
            Response.Write "<a href=""products.asp?cat=smd&link=" & lk & """ class=""cat-item"">" & nm & "<span class=""pdf-icon"">📄</span></a>"
            rSmd.MoveNext
          Loop
          rSmd.Close: Set rSmd = Nothing
          cSmd.Close: Set cSmd = Nothing
          %>
        </div>
        <div class="cat-card-footer"><a href="products.asp?cat=smd" class="cat-more">查看全部 →</a></div>
      </div>

      <!-- PUSH -->
      <div class="cat-card">
        <div class="cat-card-header">
          <span class="cat-icon">🔘</span>
          <span class="cat-title">按鈕開關 Push Switches</span>
        </div>
        <div class="cat-card-body">
          <%
          Dim cPus, rPus
          Set cPus = GetConn()
          Set rPus = cPus.Execute("SELECT TOP 6 * FROM push_s WHERE name IS NOT NULL AND name <> '' ORDER BY id")
          Do While Not rPus.EOF
            Response.Write "<a href=""products.asp?cat=push_s&link=" & rPus("link") & """ class=""cat-item"">" & rPus("name") & "<span class=""pdf-icon"">📄</span></a>"
            rPus.MoveNext
          Loop
          rPus.Close: Set rPus = Nothing
          cPus.Close: Set cPus = Nothing
          %>
        </div>
        <div class="cat-card-footer"><a href="products.asp?cat=push_s" class="cat-more">查看全部 →</a></div>
      </div>

      <!-- ROCKER -->
      <div class="cat-card">
        <div class="cat-card-header">
          <span class="cat-icon">🔄</span>
          <span class="cat-title">翹板開關 Rocker Switches</span>
        </div>
        <div class="cat-card-body">
          <%
          Dim cRoc, rRoc
          Set cRoc = GetConn()
          Set rRoc = cRoc.Execute("SELECT TOP 6 * FROM rocker_s WHERE name IS NOT NULL AND name <> '' ORDER BY id")
          Do While Not rRoc.EOF
            Response.Write "<a href=""products.asp?cat=rocker_s&link=" & rRoc("link") & """ class=""cat-item"">" & rRoc("name") & "<span class=""pdf-icon"">📄</span></a>"
            rRoc.MoveNext
          Loop
          rRoc.Close: Set rRoc = Nothing
          cRoc.Close: Set cRoc = Nothing
          %>
        </div>
        <div class="cat-card-footer"><a href="products.asp?cat=rocker_s" class="cat-more">查看全部 →</a></div>
      </div>

      <!-- SENSOR -->
      <div class="cat-card">
        <div class="cat-card-header">
          <span class="cat-icon">📡</span>
          <span class="cat-title">感應開關 Sensor Switches</span>
        </div>
        <div class="cat-card-body">
          <%
          Dim cSen, rSen
          Set cSen = GetConn()
          Set rSen = cSen.Execute("SELECT TOP 6 * FROM sensor_s WHERE name IS NOT NULL AND name <> '' ORDER BY id")
          Do While Not rSen.EOF
            Response.Write "<a href=""products.asp?cat=sensor_s&link=" & rSen("link") & """ class=""cat-item"">" & rSen("name") & "<span class=""pdf-icon"">📄</span></a>"
            rSen.MoveNext
          Loop
          rSen.Close: Set rSen = Nothing
          cSen.Close: Set cSen = Nothing
          %>
        </div>
        <div class="cat-card-footer"><a href="products.asp?cat=sensor_s" class="cat-more">查看全部 →</a></div>
      </div>

      <!-- SLIDE -->
      <div class="cat-card">
        <div class="cat-card-header">
          <span class="cat-icon">➡️</span>
          <span class="cat-title">滑動開關 Slide Switches</span>
        </div>
        <div class="cat-card-body">
          <%
          Dim cSli, rSli
          Set cSli = GetConn()
          Set rSli = cSli.Execute("SELECT TOP 6 * FROM slide_s WHERE name IS NOT NULL AND name <> '' ORDER BY id")
          Do While Not rSli.EOF
            Response.Write "<a href=""products.asp?cat=slide_s&link=" & rSli("link") & """ class=""cat-item"">" & rSli("name") & "<span class=""pdf-icon"">📄</span></a>"
            rSli.MoveNext
          Loop
          rSli.Close: Set rSli = Nothing
          cSli.Close: Set cSli = Nothing
          %>
        </div>
        <div class="cat-card-footer"><a href="products.asp?cat=slide_s" class="cat-more">查看全部 →</a></div>
      </div>

      <!-- DIP -->
      <div class="cat-card">
        <div class="cat-card-header">
          <span class="cat-icon">🔌</span>
          <span class="cat-title">DIP 開關 DIP Switches</span>
        </div>
        <div class="cat-card-body">
          <%
          Dim cDip, rDip
          Set cDip = GetConn()
          Set rDip = cDip.Execute("SELECT TOP 6 * FROM dip_s WHERE name IS NOT NULL AND name <> '' ORDER BY id")
          Do While Not rDip.EOF
            Response.Write "<a href=""products.asp?cat=dip_s&link=" & rDip("link") & """ class=""cat-item"">" & rDip("name") & "<span class=""pdf-icon"">📄</span></a>"
            rDip.MoveNext
          Loop
          rDip.Close: Set rDip = Nothing
          cDip.Close: Set cDip = Nothing
          %>
        </div>
        <div class="cat-card-footer"><a href="products.asp?cat=dip_s" class="cat-more">查看全部 →</a></div>
      </div>

      <!-- DC + EAR -->
      <div class="cat-card">
        <div class="cat-card-header">
          <span class="cat-icon">🔋</span>
          <span class="cat-title">DC 電源 / 耳機插座</span>
        </div>
        <div class="cat-card-body">
          <%
          Dim cDc, rDc
          Set cDc = GetConn()
          Set rDc = cDc.Execute("SELECT TOP 3 * FROM dc WHERE name IS NOT NULL AND name <> '' ORDER BY id")
          Do While Not rDc.EOF
            Response.Write "<a href=""products.asp?cat=dc&link=" & rDc("link") & """ class=""cat-item"">" & rDc("name") & "<span class=""pdf-icon"">📄</span></a>"
            rDc.MoveNext
          Loop
          rDc.Close: Set rDc = Nothing
          cDc.Close: Set cDc = Nothing

          Dim cEar, rEar
          Set cEar = GetConn()
          Set rEar = cEar.Execute("SELECT TOP 3 * FROM ear_j WHERE name IS NOT NULL AND name <> '' ORDER BY id")
          Do While Not rEar.EOF
            Response.Write "<a href=""products.asp?cat=ear_j&link=" & rEar("link") & """ class=""cat-item"">" & rEar("name") & "<span class=""pdf-icon"">📄</span></a>"
            rEar.MoveNext
          Loop
          rEar.Close: Set rEar = Nothing
          cEar.Close: Set cEar = Nothing
          %>
        </div>
        <div class="cat-card-footer"><a href="products.asp?cat=dc" class="cat-more">查看全部 →</a></div>
      </div>

      <!-- CONNECTORS -->
      <div class="cat-card">
        <div class="cat-card-header">
          <span class="cat-icon">🔗</span>
          <span class="cat-title">連接器 Connectors</span>
        </div>
        <div class="cat-card-body">
          <%
          Dim cCon, rCon
          Set cCon = GetConn()
          Set rCon = cCon.Execute("SELECT TOP 6 * FROM connectors WHERE name IS NOT NULL AND name <> '' ORDER BY id")
          Do While Not rCon.EOF
            Response.Write "<a href=""products.asp?cat=connectors&link=" & rCon("link") & """ class=""cat-item"">" & rCon("name") & "<span class=""pdf-icon"">📄</span></a>"
            rCon.MoveNext
          Loop
          rCon.Close: Set rCon = Nothing
          cCon.Close: Set cCon = Nothing
          %>
        </div>
        <div class="cat-card-footer"><a href="products.asp?cat=connectors" class="cat-more">查看全部 →</a></div>
      </div>

      <!-- NEW PRODUCTS -->
      <div class="cat-card" style="border:2px solid #CC0000;">
        <div class="cat-card-header" style="background:linear-gradient(135deg,#003399,#0044CC);">
          <span class="cat-icon">⭐</span>
          <span class="cat-title">新產品 New Products</span>
        </div>
        <div class="cat-card-body">
          <%
          Dim cNew, rNew
          Set cNew = GetConn()
          Set rNew = cNew.Execute("SELECT TOP 8 * FROM news WHERE name IS NOT NULL AND name <> '' ORDER BY id DESC")
          Do While Not rNew.EOF
            Response.Write "<a href=""products.asp?cat=news&link=" & rNew("link") & """ class=""cat-item"">" & rNew("name") & " <span class=""new-badge"">NEW</span><span class=""pdf-icon"">📄</span></a>"
            rNew.MoveNext
          Loop
          rNew.Close: Set rNew = Nothing
          cNew.Close: Set cNew = Nothing
          %>
        </div>
        <div class="cat-card-footer"><a href="products.asp?cat=news" class="cat-more">查看全部新品 →</a></div>
      </div>

    </div>
  </div>
</section>

<!-- ===== SEARCH CTA SECTION ===== -->
<section style="background:linear-gradient(135deg,#CC0000 0%,#990000 100%);padding:50px 0;color:#fff;">
  <div class="container" style="text-align:center;">
    <h2 style="font-family:'Rajdhani',sans-serif;font-size:36px;font-weight:700;margin-bottom:12px;">料號搜尋 Search Part No.</h2>
    <p style="font-size:16px;opacity:0.85;margin-bottom:28px;">輸入料號快速找到對應的規格圖檔 PDF</p>
    <form action="search.asp" method="get" style="display:flex;max-width:500px;margin:0 auto;gap:0;border-radius:6px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,0.2);">
      <input type="text" name="partno" placeholder="輸入料號 e.g. STS-05, PTS-056..." style="flex:1;padding:14px 20px;border:none;font-size:15px;outline:none;font-family:inherit;">
      <button type="submit" style="background:#003399;color:#fff;border:none;padding:0 28px;font-size:14px;font-weight:700;cursor:pointer;font-family:inherit;white-space:nowrap;transition:background 0.2s;" onmouseover="this.style.background='#002277'" onmouseout="this.style.background='#003399'">
        <i class="fas fa-search"></i> 搜尋
      </button>
    </form>
  </div>
</section>

<!-- ===== CATALOGS ===== -->
<section class="section section-alt">
  <div class="container">
    <div class="section-header">
      <div class="section-tag">CATALOG</div>
      <h2 class="section-title">型錄下載</h2>
      <p class="section-desc">下載最新的 HCH 產品型錄，了解完整的產品規格與資訊。</p>
      <div class="section-divider"></div>
    </div>
    <div class="catalog-grid">
      <div class="catalog-card">
        <div class="catalog-img"><i class="fas fa-file-pdf" style="font-size:56px;color:#CC0000;"></i></div>
        <h4>綜合型錄 01</h4>
        <a href="html/CON_01.pdf" target="_blank"><i class="fas fa-download"></i> 下載 PDF</a>
      </div>
      <div class="catalog-card">
        <div class="catalog-img"><i class="fas fa-file-pdf" style="font-size:56px;color:#CC0000;"></i></div>
        <h4>綜合型錄 02</h4>
        <a href="html/CON_02.pdf" target="_blank"><i class="fas fa-download"></i> 下載 PDF</a>
      </div>
      <div class="catalog-card">
        <div class="catalog-img"><i class="fas fa-file-pdf" style="font-size:56px;color:#CC0000;"></i></div>
        <h4>綜合型錄 03</h4>
        <a href="html/CON_03.pdf" target="_blank"><i class="fas fa-download"></i> 下載 PDF</a>
      </div>
      <div class="catalog-card">
        <div class="catalog-img"><i class="fas fa-file-pdf" style="font-size:56px;color:#CC0000;"></i></div>
        <h4>綜合型錄 04</h4>
        <a href="html/CON_04.pdf" target="_blank"><i class="fas fa-download"></i> 下載 PDF</a>
      </div>
    </div>
    <div style="text-align:center;margin-top:30px;">
      <a href="catalog.asp" class="btn-primary">查看所有型錄</a>
    </div>
  </div>
</section>

<!-- ===== CTA CONTACT ===== -->
<section class="section section-dark">
  <div class="container" style="text-align:center;">
    <div class="section-header">
      <div class="section-tag" style="background:rgba(204,0,0,0.3);color:#FF6666;">CONTACT US</div>
      <h2 class="section-title">需要協助？</h2>
      <p class="section-desc" style="color:rgba(255,255,255,0.6);">歡迎透過 LINE、Email 或電話與我們聯繫，專業業務人員為您服務。</p>
      <div class="section-divider"></div>
    </div>
    <div style="display:flex;gap:16px;justify-content:center;flex-wrap:wrap;margin-top:10px;">
      <a href="https://line.me/R/ti/p/@307puwgm" target="_blank" rel="noopener" class="btn-line" style="padding:13px 28px;border-radius:4px;font-size:15px;">
        <svg style="width:20px;height:20px;fill:currentColor;" viewBox="0 0 24 24"><path d="M19.952 12.375C19.952 8.33 15.9 5.03 11 5.03c-4.9 0-8.952 3.3-8.952 7.345 0 3.63 3.22 6.674 7.572 7.25.295.064.696.196.797.45.092.23.06.59.03.822 0 0-.107.639-.13.775-.04.23-.185.9.79.49.974-.408 5.254-3.094 7.17-5.297 1.32-1.452 1.673-2.904 1.673-4.49z"/></svg>
        LINE 洽詢
      </a>
      <a href="mailto:hch@hchtwn.com.tw" class="btn-primary" style="padding:13px 28px;font-size:15px;">
        <i class="fas fa-envelope" style="margin-right:6px;"></i>Email 聯絡
      </a>
      <a href="contact.asp" class="btn-outline" style="padding:13px 28px;font-size:15px;">
        <i class="fas fa-paper-plane" style="margin-right:6px;"></i>線上留言
      </a>
    </div>
  </div>
</section>

<%@include file="inc_footer.asp"%>

<%
Function FileExists(fPath)
  Dim fso
  Set fso = Server.CreateObject("Scripting.FileSystemObject")
  FileExists = fso.FileExists(Server.MapPath(fPath))
  Set fso = Nothing
End Function
%>
