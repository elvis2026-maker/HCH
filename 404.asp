<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Response.Status = "404 Not Found"
PageTitle = "找不到頁面 404"
%>
<%@include file="inc_header.asp"%>
<section class="section">
  <div class="container" style="text-align:center;padding:80px 20px;">
    <div style="font-family:'Rajdhani',sans-serif;font-size:120px;font-weight:900;color:#CC0000;line-height:1;margin-bottom:20px;">404</div>
    <h2 style="font-size:28px;font-weight:700;margin-bottom:14px;color:#111;">找不到此頁面</h2>
    <p style="font-size:16px;color:#666;margin-bottom:32px;">您所尋找的頁面不存在或已移除。<br>The page you're looking for doesn't exist.</p>
    <div style="display:flex;gap:14px;justify-content:center;flex-wrap:wrap;">
      <a href="index.asp" class="btn-primary"><i class="fas fa-home" style="margin-right:8px;"></i>返回首頁</a>
      <a href="search.asp" style="display:inline-flex;align-items:center;gap:8px;background:#003399;color:#fff;padding:13px 26px;border-radius:4px;font-weight:700;font-size:14px;">
        <i class="fas fa-search"></i> 搜尋料號
      </a>
    </div>
  </div>
</section>
<%@include file="inc_footer.asp"%>
