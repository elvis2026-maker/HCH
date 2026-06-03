<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Function GetConn()
  Dim conn, strCnn
  Set conn = Server.CreateObject("ADODB.Connection")
  strCnn = "Driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath("database/data_new.mdb")
  conn.Open strCnn
  Set GetConn = conn
End Function

' Table number to category/description mapping
Function TableName(t)
  Select Case CStr(t)
    Case "1"  : TableName = "SMD TACTILE SWITCHES"
    Case "2"  : TableName = "PUSH SWITCHES / MINIATURE SWITCHES"
    Case "3"  : TableName = "ROCKER SWITCHES"
    Case "4"  : TableName = "SENSOR SWITCHES / MICRO SWITCHES"
    Case "5"  : TableName = "SLIDE SWITCHES"
    Case "6"  : TableName = "DIP SWITCHES / FUNCTION SWITCHES"
    Case "7"  : TableName = "DC POWER JACKS"
    Case "8"  : TableName = "EARPHONE JACKS"
    Case "9"  : TableName = "POTENTIOMETERS / ENCODERS"
    Case "10" : TableName = "CONNECTORS / LENSES / USB / FPC / ESD"
    Case Else : TableName = "ELECTRONIC COMPONENTS"
  End Select
End Function

PageTitle = "料號搜尋 Search Part No."

Dim partno, hasSearch, results()
Dim resultCount : resultCount = 0
partno = Trim(Request.QueryString("partno"))
hasSearch = (partno <> "")

If hasSearch Then
  ' Sanitize input
  Dim cleanPart : cleanPart = ""
  Dim i
  For i = 1 To Len(partno)
    Dim c : c = Mid(partno, i, 1)
    If c >= "A" And c <= "Z" Or c >= "a" And c <= "z" Or c >= "0" And c <= "9" Or c = "-" Or c = "." Or c = "_" Or c = " " Then
      cleanPart = cleanPart & c
    End If
  Next
  partno = cleanPart

  If partno <> "" Then
    Dim connS, rsS
    Set connS = GetConn()
    Dim sqlS : sqlS = "SELECT * FROM search WHERE NAME LIKE '%" & Replace(partno,"'","") & "%' ORDER BY id"
    Set rsS = connS.Execute(sqlS)
    Do While Not rsS.EOF
      ReDim Preserve results(resultCount)
      results(resultCount) = rsS("NAME") & "|" & rsS("LINK") & "|" & rsS("TABLE")
      resultCount = resultCount + 1
      rsS.MoveNext
    Loop
    rsS.Close: Set rsS = Nothing
    connS.Close: Set connS = Nothing
  End If
End If
%>
<%@include file="inc_header.asp"%>

<!-- BREADCRUMB -->
<div class="breadcrumb">
  <div class="container">
    <ul class="breadcrumb-list">
      <li><a href="index.asp">首頁</a></li>
      <li>料號搜尋</li>
    </ul>
  </div>
</div>

<!-- PAGE HERO -->
<div class="page-hero">
  <div class="container">
    <h1><i class="fas fa-search" style="font-size:32px;margin-right:14px;opacity:0.8;"></i>料號搜尋 Search Part No.</h1>
    <p>輸入料號關鍵字可搜尋對應的 PDF 規格圖檔。本資料庫收錄超過 2,900 筆料號。</p>
  </div>
</div>

<section class="section">
  <div class="container" style="max-width:900px;">

    <!-- SEARCH FORM -->
    <div style="background:#fff;border:1px solid #E0E0E0;border-radius:10px;padding:36px;box-shadow:0 4px 20px rgba(0,0,0,0.08);margin-bottom:36px;">
      <h2 style="font-family:'Rajdhani',sans-serif;font-size:26px;font-weight:700;margin-bottom:8px;color:#111;">
        <i class="fas fa-search" style="color:#CC0000;margin-right:10px;"></i>搜尋料號
      </h2>
      <p style="font-size:14px;color:#666;margin-bottom:24px;">
        在下方輸入完整或部分料號，系統將搜尋所有符合的品項並提供 PDF 規格圖檔連結。<br>
        Enter full or partial part number to search for matching items and PDF datasheets.
      </p>
      <form action="search.asp" method="get" id="searchForm">
        <div style="display:flex;gap:0;border-radius:6px;overflow:hidden;border:2px solid #E0E0E0;transition:border-color 0.2s;" id="searchWrap">
          <input type="text" name="partno" id="partnoInput" value="<%=Server.HTMLEncode(partno)%>"
            placeholder="例如: STS-05, PTS-056, RTS-12, SMD_01..."
            style="flex:1;padding:14px 18px;border:none;outline:none;font-size:15px;font-family:inherit;color:#111;"
            onfocus="document.getElementById('searchWrap').style.borderColor='#CC0000'"
            onblur="document.getElementById('searchWrap').style.borderColor='#E0E0E0'"
            autocomplete="off">
          <button type="submit" style="background:#CC0000;color:#fff;border:none;padding:0 32px;font-size:15px;font-weight:700;cursor:pointer;font-family:inherit;transition:background 0.2s;white-space:nowrap;" onmouseover="this.style.background='#990000'" onmouseout="this.style.background='#CC0000'">
            <i class="fas fa-search"></i> 搜尋
          </button>
        </div>
      </form>

      <!-- HINTS -->
      <div style="margin-top:16px;display:flex;flex-wrap:wrap;gap:8px;">
        <span style="font-size:12px;color:#999;">快速範例：</span>
        <%
        Dim hints : hints = Array("STS-05","PTS-056","ROC","RTS","DIP","DC","PJK","POT","CON","USB","FPC","ESD")
        Dim h
        For Each h In hints
          Response.Write "<a href=""search.asp?partno=" & h & """ style=""font-size:12px;padding:3px 10px;background:#F5F5F5;border:1px solid #E0E0E0;border-radius:20px;color:#333;transition:all 0.15s;"" onmouseover=""this.style.background='#FFF0F0';this.style.borderColor='#CC0000';this.style.color='#CC0000'"" onmouseout=""this.style.background='#F5F5F5';this.style.borderColor='#E0E0E0';this.style.color='#333'"">" & h & "</a>"
        Next
        %>
      </div>
    </div>

    <!-- SEARCH RESULTS -->
    <%
    If hasSearch And partno <> "" Then
      If resultCount = 0 Then
        Response.Write "<div style=""background:#fff;border:1px solid #E0E0E0;border-radius:10px;padding:48px;text-align:center;box-shadow:0 2px 10px rgba(0,0,0,0.06);"">"
        Response.Write "<i class=""fas fa-search-minus"" style=""font-size:48px;color:#ccc;margin-bottom:16px;""></i>"
        Response.Write "<h3 style=""font-size:20px;font-weight:700;margin-bottom:8px;color:#333;"">找不到符合的料號</h3>"
        Response.Write "<p style=""color:#888;font-size:14px;margin-bottom:20px;"">未找到包含「" & Server.HTMLEncode(partno) & "」的料號。請確認料號輸入是否正確，或嘗試部分關鍵字搜尋。</p>"
        Response.Write "<p style=""color:#888;font-size:13px;"">No parts found for &quot;" & Server.HTMLEncode(partno) & "&quot;. Please check the part number or try a partial search.</p>"
        Response.Write "<a href=""contact.asp"" style=""display:inline-block;margin-top:20px;background:#003399;color:#fff;padding:10px 24px;border-radius:4px;font-weight:700;font-size:14px;""><i class=""fas fa-envelope"" style=""margin-right:6px;""></i>詢問我們</a>"
        Response.Write "</div>"
      Else
        Response.Write "<div style=""background:#fff;border:1px solid #E0E0E0;border-radius:10px;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,0.08);"">"
        Response.Write "<div style=""background:linear-gradient(135deg,#990000,#CC0000);color:#fff;padding:16px 24px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px;"">"
        Response.Write "<div><i class=""fas fa-list"" style=""margin-right:8px;""></i>搜尋結果 Search Results</div>"
        Response.Write "<div style=""background:rgba(255,255,255,0.2);padding:4px 14px;border-radius:20px;font-size:13px;"">找到 <strong>" & resultCount & "</strong> 筆 / 料號「" & Server.HTMLEncode(partno) & "」</div>"
        Response.Write "</div>"
        Response.Write "<div style=""overflow-x:auto;"">"
        Response.Write "<table style=""width:100%;border-collapse:collapse;font-size:14px;"">"
        Response.Write "<thead><tr style=""background:#F5F5F5;"">"
        Response.Write "<th style=""padding:11px 16px;text-align:left;font-size:12px;font-weight:700;color:#666;letter-spacing:0.06em;border-bottom:1px solid #E0E0E0;"">#</th>"
        Response.Write "<th style=""padding:11px 16px;text-align:left;font-size:12px;font-weight:700;color:#666;letter-spacing:0.06em;border-bottom:1px solid #E0E0E0;"">料號 Part No.</th>"
        Response.Write "<th style=""padding:11px 16px;text-align:left;font-size:12px;font-weight:700;color:#666;letter-spacing:0.06em;border-bottom:1px solid #E0E0E0;"">產品類別 Category</th>"
        Response.Write "<th style=""padding:11px 16px;text-align:left;font-size:12px;font-weight:700;color:#666;letter-spacing:0.06em;border-bottom:1px solid #E0E0E0;"">規格圖檔 PDF</th>"
        Response.Write "</thead><tbody>"
        Dim r, parts
        For r = 0 To resultCount - 1
          parts = Split(results(r), "|")
          Dim pName2 : pName2 = parts(0)
          Dim pLink2 : pLink2 = parts(1)
          Dim pTable : pTable = parts(2)
          Response.Write "<tr onmouseover=""this.style.background='#FFF5F5'"" onmouseout=""this.style.background=''"">"
          Response.Write "<td style=""padding:10px 16px;color:#bbb;font-size:12px;border-bottom:1px solid #F0F0F0;"">" & (r+1) & "</td>"
          Response.Write "<td style=""padding:10px 16px;border-bottom:1px solid #F0F0F0;font-weight:600;font-family:'Rajdhani',sans-serif;font-size:15px;color:#111;"">" & Server.HTMLEncode(pName2) & "</td>"
          Response.Write "<td style=""padding:10px 16px;border-bottom:1px solid #F0F0F0;font-size:12px;color:#666;"">" & TableName(pTable) & "</td>"
          Response.Write "<td style=""padding:10px 16px;border-bottom:1px solid #F0F0F0;"">"
          If pLink2 <> "" Then
            Response.Write "<a href=""html/single/" & pLink2 & """ target=""_blank"" class=""pdf-link""><i class=""fas fa-file-pdf""></i> PDF</a>"
          Else
            Response.Write "<span style=""color:#ccc;font-size:12px;"">—</span>"
          End If
          Response.Write "</td></tr>"
        Next
        Response.Write "</tbody></table></div></div>"
      End If
    ElseIf hasSearch And partno = "" Then
      Response.Write "<div style=""background:#FFF3CD;border:1px solid #FFE69C;border-radius:8px;padding:16px 20px;color:#856404;""><i class=""fas fa-exclamation-triangle"" style=""margin-right:8px;""></i>請輸入料號後再搜尋。Please enter a part number.</div>"
    End If
    %>

    <!-- INFO BOX -->
    <div style="margin-top:36px;background:#F0F4FF;border:1px solid #C7D7FF;border-radius:10px;padding:24px 28px;">
      <h4 style="font-size:16px;font-weight:700;color:#003399;margin-bottom:10px;"><i class="fas fa-info-circle" style="margin-right:8px;"></i>搜尋說明</h4>
      <ul style="font-size:13px;color:#444;line-height:2;">
        <li>• 本資料庫收錄超過 <strong>2,900 筆</strong>料號，涵蓋所有 HCH 產品系列。</li>
        <li>• 可輸入完整料號（如 STS-05-A）或部分料號（如 STS）進行搜尋。</li>
        <li>• 搜尋結果提供 PDF 規格圖檔直接下載連結。</li>
        <li>• 若找不到所需料號，請透過 <a href="contact.asp" style="color:#CC0000;font-weight:600;">聯絡我們</a> 或 <a href="https://line.me/R/ti/p/@307puwgm" target="_blank" style="color:#06C755;font-weight:600;">LINE</a> 詢問。</li>
        <li>• This database contains 2,900+ part numbers. Search by full or partial part number.</li>
      </ul>
    </div>
  </div>
</section>

<%@include file="inc_footer.asp"%>
