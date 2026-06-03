<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
' ===== DATABASE CONNECTION HELPER =====
Function GetConn()
  Dim conn, strCnn
  Set conn = Server.CreateObject("ADODB.Connection")
  strCnn = "Driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath("database/data_new.mdb")
  conn.Open strCnn
  Set GetConn = conn
End Function
%>
<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<title><%=PageTitle%> | HCH 鴻琪股份有限公司</title>
<meta name="description" content="鴻琪股份有限公司 - 用心創新第一，品質服務優先。專業製造各類開關、連接器、感應器等電子零件，ISO/SGS認證，外銷全球。">
<meta name="keywords" content="HCH,鴻琪,鴻琪股份有限公司,switches,push switches,rocker switches,sensor switches,slide switches,DIP switches,DC power jacks,earphone jacks,potentiometers,connectors,SMD,開關,電子零件,ISO,SGS">
<meta name="author" content="HCH 鴻琪股份有限公司">
<meta name="copyright" content="Copyright 1984 HCH Co. All rights reserved.">
<meta name="robots" content="index, follow">
<meta property="og:title" content="HCH 鴻琪股份有限公司">
<meta property="og:description" content="專業製造各類電子開關、連接器、感應器，ISO/SGS認證，品質保證。">
<meta property="og:url" content="https://www.hchtwn.com.tw">
<meta property="og:type" content="website">
<link rel="shortcut icon" href="images/hchfavicon.ico">
<link rel="apple-touch-icon" href="images/hchicon.png">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Rajdhani:wght@400;500;600;700&family=Noto+Sans+TC:wght@300;400;500;700&display=swap" rel="stylesheet">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
<link rel="stylesheet" href="css/style.css">
</head>
<body>

<!-- TOP BAR -->
<div class="topbar">
  <div class="container">
    <div class="topbar-contact">
      <span><i class="fas fa-phone"></i>+886-2-2393-4278</span>
      <span><i class="fas fa-envelope"></i>hch@hchtwn.com.tw</span>
      <span><i class="fas fa-map-marker-alt"></i>台灣台北市</span>
    </div>
    <div style="display:flex;align-items:center;gap:14px;">
      <span style="font-size:11px;color:#666;">成立於 1984 &nbsp;|&nbsp; ISO 認證 &nbsp;|&nbsp; SGS 認證</span>
      <div class="topbar-lang">
        <a href="index.asp">中文</a>
        <a href="index.asp">EN</a>
      </div>
    </div>
  </div>
</div>

<!-- SITE HEADER -->
<header class="site-header">
  <div class="container">
    <div class="header-inner">
      <!-- LOGO -->
      <a href="index.asp" class="logo">
        <div class="logo-text">HCH</div>
        <div class="logo-sub">
          <span>鴻琪股份有限公司</span>
          <span>HORNG CHIH HO CO., LTD.</span>
        </div>
      </a>

      <!-- SEARCH BAR -->
      <div class="header-search">
        <form class="search-wrap" action="search.asp" method="get" onsubmit="return validateSearch()">
          <input type="text" name="partno" id="searchInput" placeholder="搜尋料號 Search Part No. (e.g. STS-05)" autocomplete="off" value="<%=Request.QueryString("partno")%>">
          <button type="submit"><i class="fas fa-search"></i> 搜尋</button>
        </form>
      </div>

      <!-- ACTION BUTTONS -->
      <div class="header-actions">
        <a href="https://line.me/R/ti/p/@307puwgm" target="_blank" class="btn-line" rel="noopener">
          <svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path d="M19.952 12.375C19.952 8.33 15.9 5.03 11 5.03c-4.9 0-8.952 3.3-8.952 7.345 0 3.63 3.22 6.674 7.572 7.25.295.064.696.196.797.45.092.23.06.59.03.822 0 0-.107.639-.13.775-.04.23-.185.9.79.49.974-.408 5.254-3.094 7.17-5.297 1.32-1.452 1.673-2.904 1.673-4.49z"/></svg>
          LINE
        </a>
        <a href="contact.asp" class="btn-contact">
          <i class="fas fa-envelope"></i> 聯絡我們
        </a>
      </div>
    </div>
  </div>
</header>

<!-- MAIN NAVIGATION -->
<nav class="main-nav">
  <div class="container">
    <div class="nav-toggle-wrap" style="display:flex;justify-content:space-between;align-items:center;">
      <span style="color:rgba(255,255,255,0.7);font-size:13px;display:none;" class="nav-label-mobile">產品選單</span>
      <button class="nav-toggle" id="navToggle" aria-label="選單">
        <span></span><span></span><span></span>
      </button>
    </div>
    <ul class="nav-list" id="navList">
      <li class="nav-item"><a href="index.asp" class="nav-link"><i class="fas fa-home" style="margin-right:4px;font-size:12px;"></i>首頁</a></li>
      <li class="nav-item"><a href="about.asp" class="nav-link">公司介紹</a></li>
      <li class="nav-item">
        <a href="#" class="nav-link">SMD 觸控開關 <span class="arrow">▾</span></a>
        <div class="dropdown-menu">
          <%
          Dim conn1, rs1
          Set conn1 = GetConn()
          Dim sql1 : sql1 = "SELECT * FROM smd WHERE name IS NOT NULL AND name <> '' ORDER BY id"
          Set rs1 = conn1.Execute(sql1)
          Do While Not rs1.EOF
            Response.Write "<a href=""products.asp?cat=smd&link=" & rs1("link") & """ class=""dropdown-item"">" & rs1("name") & "</a>"
            rs1.MoveNext
          Loop
          rs1.Close: Set rs1 = Nothing
          conn1.Close: Set conn1 = Nothing
          %>
        </div>
      </li>
      <li class="nav-item">
        <a href="#" class="nav-link">按鈕開關 <span class="arrow">▾</span></a>
        <div class="dropdown-menu">
          <%
          Dim conn2, rs2
          Set conn2 = GetConn()
          Set rs2 = conn2.Execute("SELECT * FROM push_s WHERE name IS NOT NULL AND name <> '' ORDER BY id")
          Do While Not rs2.EOF
            Response.Write "<a href=""products.asp?cat=push_s&link=" & rs2("link") & """ class=""dropdown-item"">" & rs2("name") & "</a>"
            rs2.MoveNext
          Loop
          rs2.Close: Set rs2 = Nothing
          conn2.Close: Set conn2 = Nothing
          %>
        </div>
      </li>
      <li class="nav-item">
        <a href="#" class="nav-link">翹板開關 <span class="arrow">▾</span></a>
        <div class="dropdown-menu">
          <%
          Dim conn3, rs3
          Set conn3 = GetConn()
          Set rs3 = conn3.Execute("SELECT * FROM rocker_s WHERE name IS NOT NULL AND name <> '' ORDER BY id")
          Do While Not rs3.EOF
            Response.Write "<a href=""products.asp?cat=rocker_s&link=" & rs3("link") & """ class=""dropdown-item"">" & rs3("name") & "</a>"
            rs3.MoveNext
          Loop
          rs3.Close: Set rs3 = Nothing
          conn3.Close: Set conn3 = Nothing
          %>
        </div>
      </li>
      <li class="nav-item">
        <a href="#" class="nav-link">更多產品 <span class="arrow">▾</span></a>
        <div class="dropdown-menu">
          <span class="dropdown-label">感應 / 滑動 / DIP 開關</span>
          <a href="products.asp?cat=sensor_s" class="dropdown-item">SENSOR SWITCHES</a>
          <a href="products.asp?cat=slide_s" class="dropdown-item">SLIDE SWITCHES</a>
          <a href="products.asp?cat=dip_s" class="dropdown-item">DIP SWITCHES</a>
          <span class="dropdown-label">電源插座 / 音頻</span>
          <a href="products.asp?cat=dc" class="dropdown-item">DC POWER JACKS</a>
          <a href="products.asp?cat=ear_j" class="dropdown-item">EARPHONE JACKS</a>
          <span class="dropdown-label">電位計 / 連接器</span>
          <a href="products.asp?cat=potentio" class="dropdown-item">POTENTIOMETERS / ENCODERS</a>
          <a href="products.asp?cat=connectors" class="dropdown-item">CONNECTORS / LENSES</a>
          <a href="products.asp?cat=conn_pse" class="dropdown-item">USB CONNECTORS</a>
          <a href="products.asp?cat=conn_fus" class="dropdown-item">ESD / FUSE / MOTOR / PPTC</a>
          <a href="products.asp?cat=conn_fpc" class="dropdown-item">FPC.FFC CONNECTORS</a>
        </div>
      </li>
      <li class="nav-item"><a href="products.asp?cat=news" class="nav-link" style="background:rgba(204,0,0,0.3);"><i class="fas fa-star" style="margin-right:4px;font-size:10px;color:#FFD700;"></i>新產品</a></li>
      <li class="nav-item"><a href="catalog.asp" class="nav-link">型錄下載</a></li>
      <li class="nav-item"><a href="search.asp" class="nav-link"><i class="fas fa-search" style="margin-right:4px;font-size:11px;"></i>料號搜尋</a></li>
      <li class="nav-item"><a href="contact.asp" class="nav-link">聯絡我們</a></li>
    </ul>
  </div>
</nav>

<script>
function validateSearch(){
  var v = document.getElementById('searchInput').value.trim();
  if(!v){ alert('請輸入料號'); return false; }
  return true;
}
// Mobile Nav Toggle
document.addEventListener('DOMContentLoaded', function(){
  var toggle = document.getElementById('navToggle');
  var navList = document.getElementById('navList');
  if(toggle){ toggle.addEventListener('click', function(){ navList.classList.toggle('open'); }); }
  // Sub-menu on mobile
  var navItems = document.querySelectorAll('.nav-item');
  navItems.forEach(function(item){
    var link = item.querySelector('.nav-link');
    var menu = item.querySelector('.dropdown-menu');
    if(menu && link){
      link.addEventListener('click', function(e){
        if(window.innerWidth <= 768){
          e.preventDefault();
          item.classList.toggle('open');
        }
      });
    }
  });
});
</script>
