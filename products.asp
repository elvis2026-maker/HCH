<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Function GetConn()
  Dim conn, strCnn
  Set conn = Server.CreateObject("ADODB.Connection")
  strCnn = "Driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath("database/data_new.mdb")
  conn.Open strCnn
  Set GetConn = conn
End Function

' Get parameters
Dim cat, lnk
cat = LCase(Trim(Request.QueryString("cat")))
lnk = Trim(Request.QueryString("link"))

' Category display names
Function CatName(c)
  Select Case c
    Case "smd"        : CatName = "SMD 觸控開關 Tactile Switches"
    Case "push_s"     : CatName = "按鈕開關 Push Switches"
    Case "rocker_s"   : CatName = "翹板開關 Rocker Switches"
    Case "sensor_s"   : CatName = "感應開關 Sensor Switches"
    Case "slide_s"    : CatName = "滑動開關 Slide Switches"
    Case "dip_s"      : CatName = "DIP 開關 DIP Switches"
    Case "dc"         : CatName = "DC 電源插座 DC Power Jacks"
    Case "ear_j"      : CatName = "耳機插座 Earphone Jacks"
    Case "potentio"   : CatName = "電位計 / 編碼器 Potentiometers"
    Case "connectors" : CatName = "連接器 / 鏡片 Connectors"
    Case "conn_pse"   : CatName = "USB 連接器 USB Connectors"
    Case "conn_fus"   : CatName = "ESD / 保險絲 / 馬達 / PPTC"
    Case "conn_fpc"   : CatName = "FPC.FFC 連接器"
    Case "news"       : CatName = "新產品 New Products"
    Case Else         : CatName = "產品列表"
  End Select
End Function

' Validate table name to prevent SQL injection
Function ValidCat(c)
  Dim valid(13)
  valid(0) = "smd": valid(1) = "push_s": valid(2) = "rocker_s"
  valid(3) = "sensor_s": valid(4) = "slide_s": valid(5) = "dip_s"
  valid(6) = "dc": valid(7) = "ear_j": valid(8) = "potentio"
  valid(9) = "connectors": valid(10) = "conn_pse": valid(11) = "conn_fus"
  valid(12) = "conn_fpc": valid(13) = "news"
  Dim i
  For i = 0 To 13
    If valid(i) = c Then ValidCat = True: Exit Function
  Next
  ValidCat = False
End Function

PageTitle = CatName(cat)
%>
<%@include file="inc_header.asp"%>

<!-- BREADCRUMB -->
<div class="breadcrumb">
  <div class="container">
    <ul class="breadcrumb-list">
      <li><a href="index.asp">首頁</a></li>
      <li><a href="products.asp?cat=smd">產品目錄</a></li>
      <li><%=CatName(cat)%></li>
    </ul>
  </div>
</div>

<!-- PAGE HERO -->
<div class="page-hero">
  <div class="container">
    <h1><i class="fas fa-th-list" style="font-size:32px;margin-right:14px;opacity:0.8;"></i><%=CatName(cat)%></h1>
    <p>點擊下方料號群組可查看對應的 PDF 規格圖檔。Click a series to view the PDF datasheet.</p>
  </div>
</div>

<section class="section">
  <div class="container">
    <div style="display:grid;grid-template-columns:280px 1fr;gap:30px;align-items:start;">

      <!-- SIDEBAR CATEGORIES -->
      <aside>
        <div style="background:#fff;border:1px solid #E0E0E0;border-radius:8px;overflow:hidden;box-shadow:0 2px 10px rgba(0,0,0,0.06);">
          <div style="background:linear-gradient(135deg,#990000,#CC0000);color:#fff;padding:14px 18px;font-weight:700;font-size:14px;letter-spacing:0.04em;">
            <i class="fas fa-sitemap" style="margin-right:8px;"></i>全部產品分類
          </div>
          <%
          Dim cats(13), catNames(13)
          cats(0) = "smd":        catNames(0) = "SMD 觸控開關"
          cats(1) = "push_s":     catNames(1) = "按鈕開關 Push Switches"
          cats(2) = "rocker_s":   catNames(2) = "翹板開關 Rocker Switches"
          cats(3) = "sensor_s":   catNames(3) = "感應開關 Sensor Switches"
          cats(4) = "slide_s":    catNames(4) = "滑動開關 Slide Switches"
          cats(5) = "dip_s":      catNames(5) = "DIP 開關"
          cats(6) = "dc":         catNames(6) = "DC 電源插座"
          cats(7) = "ear_j":      catNames(7) = "耳機插座 Earphone Jacks"
          cats(8) = "potentio":   catNames(8) = "電位計 / 編碼器"
          cats(9) = "connectors": catNames(9) = "連接器 / 鏡片"
          cats(10) = "conn_pse":  catNames(10) = "USB 連接器"
          cats(11) = "conn_fus":  catNames(11) = "ESD / 保險絲 / 馬達"
          cats(12) = "conn_fpc":  catNames(12) = "FPC.FFC 連接器"
          cats(13) = "news":      catNames(13) = "⭐ 新產品 NEW"
          Dim ci
          For ci = 0 To 13
            Dim isActive : If cats(ci) = cat Then isActive = " style=""background:#FFF5F5;border-left-color:#CC0000;color:#CC0000;font-weight:700;""" Else isActive = ""
            Response.Write "<a href=""products.asp?cat=" & cats(ci) & """ style=""display:block;padding:9px 18px;font-size:13px;color:#333;border-left:3px solid transparent;transition:all 0.15s;"" " & isActive & " onmouseover=""if(this.style.color!='rgb(204,0,0)')this.style.background='#F8F8F8'"" onmouseout=""if(this.style.color!='rgb(204,0,0)')this.style.background=''"" >" & catNames(ci) & "</a>"
          Next
          %>
        </div>

        <!-- QUICK SEARCH -->
        <div style="background:#fff;border:1px solid #E0E0E0;border-radius:8px;overflow:hidden;margin-top:20px;box-shadow:0 2px 10px rgba(0,0,0,0.06);">
          <div style="background:linear-gradient(135deg,#002277,#003399);color:#fff;padding:14px 18px;font-weight:700;font-size:14px;">
            <i class="fas fa-search" style="margin-right:8px;"></i>料號快速搜尋
          </div>
          <div style="padding:16px;">
            <form action="search.asp" method="get">
              <input type="text" name="partno" placeholder="輸入料號..." style="width:100%;padding:9px 12px;border:1.5px solid #E0E0E0;border-radius:4px;font-size:13px;font-family:inherit;outline:none;margin-bottom:10px;">
              <button type="submit" style="width:100%;background:#CC0000;color:#fff;border:none;padding:10px;border-radius:4px;font-size:13px;font-weight:700;cursor:pointer;font-family:inherit;">
                <i class="fas fa-search"></i> 搜尋料號
              </button>
            </form>
          </div>
        </div>
      </aside>

      <!-- MAIN CONTENT -->
      <main>
        <%
        If cat = "" Or Not ValidCat(cat) Then
          Response.Write "<div style=""text-align:center;padding:60px;color:#666;""><i class=""fas fa-exclamation-triangle"" style=""font-size:48px;color:#CC0000;margin-bottom:16px;""></i><br><h3>請選擇左側的產品分類</h3></div>"
        Else
          ' Show PDF viewer if link is provided
          If lnk <> "" Then
            ' Sanitize link (alphanumeric, hyphen, underscore only)
            Dim cleanLnk : cleanLnk = ""
            Dim charSet : charSet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_"
            Dim ci2
            For ci2 = 1 To Len(lnk)
              If InStr(charSet, Mid(lnk, ci2, 1)) > 0 Then cleanLnk = cleanLnk & Mid(lnk, ci2, 1)
            Next
            lnk = cleanLnk
            Response.Write "<div style=""margin-bottom:20px;"">"
            Response.Write "<a href=""products.asp?cat=" & cat & """ style=""display:inline-flex;align-items:center;gap:6px;color:#003399;font-size:13px;font-weight:600;margin-bottom:16px;"">← 返回列表</a>"
            Response.Write "<div style=""background:#fff;border:1px solid #E0E0E0;border-radius:8px;overflow:hidden;box-shadow:0 2px 10px rgba(0,0,0,0.08);"">"
            Response.Write "<div style=""background:linear-gradient(135deg,#990000,#CC0000);color:#fff;padding:14px 20px;font-weight:700;"">"
            Response.Write "<i class=""fas fa-file-pdf"" style=""margin-right:8px;""></i>" & lnk & " - 規格圖檔</div>"
            Response.Write "<div style=""padding:20px;"">"
            Response.Write "<p style=""margin-bottom:14px;font-size:13px;color:#666;"">PDF 規格圖檔如未顯示，請點擊下方按鈕直接開啟。</p>"
            Response.Write "<a href=""html/single/" & lnk & ".pdf"" target=""_blank"" style=""display:inline-flex;align-items:center;gap:6px;background:#CC0000;color:#fff;padding:10px 20px;border-radius:4px;font-weight:700;font-size:14px;margin-bottom:16px;""><i class=""fas fa-external-link-alt""></i> 開啟 PDF 圖檔</a>"
            Response.Write "<br><iframe src=""html/single/" & lnk & ".pdf"" style=""width:100%;height:680px;border:1px solid #eee;border-radius:4px;"" title=""" & lnk & """>您的瀏覽器不支援 PDF 內嵌顯示，請點上方按鈕開啟。</iframe>"
            Response.Write "</div></div></div>"
          End If

          ' List all items in category
          Dim connCat, rsCat
          Set connCat = GetConn()
          Dim sqlCat : sqlCat = "SELECT * FROM " & cat & " WHERE name IS NOT NULL AND name <> '' ORDER BY id"
          Set rsCat = connCat.Execute(sqlCat)

          Response.Write "<div class=""product-table-wrap"">"
          Response.Write "<table class=""product-table"">"
          Response.Write "<thead><tr><th style=""width:50px;"">#</th><th>產品系列名稱 Product Series</th><th style=""width:180px;"">規格圖檔 PDF</th></tr></thead>"
          Response.Write "<tbody>"
          Dim rowNum : rowNum = 0
          Do While Not rsCat.EOF
            rowNum = rowNum + 1
            Dim pName : pName = rsCat("name")
            Dim pLink : pLink = rsCat("link")
            Dim isNew : isNew = (InStr(LCase(pName), "new") > 0)
            Response.Write "<tr>"
            Response.Write "<td style=""color:#999;font-size:12px;"">" & rowNum & "</td>"
            Response.Write "<td><a href=""products.asp?cat=" & cat & "&link=" & pLink & """ style=""color:#111;font-weight:500;"">"
            Response.Write pName
            If isNew Then Response.Write " <span class=""new-badge"">NEW</span>"
            Response.Write "</a></td>"
            Response.Write "<td>"
            If pLink <> "" Then
              Response.Write "<a href=""html/single/" & pLink & ".pdf"" target=""_blank"" class=""pdf-link""><i class=""fas fa-file-pdf""></i> PDF</a>"
              Response.Write " &nbsp; <a href=""products.asp?cat=" & cat & "&link=" & pLink & """ style=""font-size:12px;color:#003399;""><i class=""fas fa-eye""></i> 預覽</a>"
            End If
            Response.Write "</td></tr>"
            rsCat.MoveNext
          Loop
          Response.Write "</tbody></table></div>"
          rsCat.Close: Set rsCat = Nothing
          connCat.Close: Set connCat = Nothing
        End If
        %>
      </main>
    </div>
  </div>
</section>

<!-- RESPONSIVE SIDEBAR TOGGLE -->
<style>
@media(max-width:768px){
  .products-layout { grid-template-columns: 1fr !important; }
}
</style>

<%@include file="inc_footer.asp"%>
