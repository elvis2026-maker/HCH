<% Option Explicit
dim objConn,strCnn,rs
set objConn = Server.CreateObject("Adodb.connection")
strCnn = "Driver={Microsoft Access Driver (*.mdb)};dbq=" & _
Server.MapPath("database/data_new.mdb")
objConn.Open strCnn
sub main()
dim choose
	choose=request("do1")
	select case choose
		case ""
			main_select
		case "add_select"
			add_choose
		case "count1"
			calculate
	end select
end sub

sub main_select()
Dim Sqlstr,a,b,c

  response.Write("<TABLE WIDTH=100% BORDER=0 CELLPADDING=0 CELLSPACING=0>" & vbCRLF)
    response.Write("<TR><TD width=492> <IMG SRC='images/data_01.gif' WIDTH=492 HEIGHT=50 border=0></TD>" & vbCRLF) 
    response.Write("<TD ROWSPAN=2 class=bg01>&nbsp; </TD>" & vbCRLF) 
  response.Write("</TR>" & vbCRLF) 
  response.Write("<TR> " & vbCRLF) 
    response.Write("<TD width=492 height=99 background='images/new_data_bg.gif'>" & vbCRLF)
	response.Write("<table width=400 border=0 align=center cellpadding=0 cellspacing=0>" & vbCRLF)
' 8 page
response.Write("<form name=form1 method=post action=manager.asp >" & vbCRLF)
	    response.Write("<tr>" & vbCRLF)
		  response.Write("<td width=99 valign=middle><div align=right>&nbsp;</div>"& vbCRLF)
      	  response.Write("<td width=168 valign=middle><div align=right>"& vbCRLF)

		response.write("<td>&nbsp;</td>")
		response.write("<td width=34 align=bottom><a href=manager.asp?do1=add_select> <img src=images/data_09.gif width=59 height=24 border=0 align=bottom> </a> </td>" & vbCRLF)
        response.Write("</div></td>" & vbCRLF)
		response.write("</tr>" & vbCRLF)
		response.Write("</form>" & vbCRLF)
'single page
response.Write("<form name=form1 method=post action=manager.asp >" & vbCRLF)		
			    response.Write("<tr>" & vbCRLF)
		  response.Write("<td width=99 valign=middle><div align=right><font color=#FFFFFF size=2>單一頁 : &nbsp;<div align=right></font></div>"& vbCRLF)
      	  response.Write("<td width=168 valign=middle><div align=right>"& vbCRLF)
		  response.Write("<select name=menu81>" & vbCRLF)
          response.Write("<option value=a selected>&nbsp;</option>" & vbCRLF)
          response.Write("<option value=smd>SMD TACTILE SWITCHES</option>" & vbCRLF)
          response.Write("<option value=function_s>FUNCTION SWITCHES</option>" & vbCRLF)
          response.Write("<option value=push_s>PUSH SWITCHES</option>" & vbCRLF)
          response.Write("<option value=sensor_s>SENSOR SWITCHES</option>" & vbCRLF)
          response.Write("<option value=rocker_s>ROCKER SWITCHES</option>" & vbCRLF)
          response.Write("<option value=slide_s>SLIDE SWITCHES</option>" & vbCRLF)
          response.Write("<option value=dip_s>DIP SWITCHES</option>" & vbCRLF)
          response.Write("<option value=ear_j>EARPHONE JACKS</option>" & vbCRLF)
          response.Write("<option value=dc>DC POWER JACKS</option>" & vbCRLF)
          response.Write("<option value=potentio>POTENTIOMETERS</option>" & vbCRLF)
          response.Write("<option value='connectors'>CONNECTORS</option>" & vbCRLF)
		  response.Write("<option value='new'>NEW PRODUCTS</option>" & vbCRLF)		  
          response.Write("</select>" & vbCRLF)
		response.write("<td><button name=submit type=submit width=34 height=24 class=input><img src=images/data_07.gif width=34 height=24 border=0></button></td>")
        response.Write("</div></td>" & vbCRLF)
		response.write("</tr>" & vbCRLF)
		response.Write("</form>" & vbCRLF)
'search
response.Write("<form name=form1 method=post action=manager.asp >" & vbCRLF)
		response.Write("<tr>" & vbCRLF) 
		  response.Write("<td width=99 valign=middle><div align=right><font color=#FFFFFF size=2>&nbsp;請輸入型號:&nbsp;</font><div align=right></div>"& vbCRLF)
      	  response.Write("<td width=168 valign=middle><div align=right>"& vbCRLF)
          response.Write("<input type=text name=search id=search >" & vbCRLF)
		response.write("<td><button name=submit type=submit width=34 height=24 class=input><img src=images/data_07.gif width=34 height=24 border=0></button></td>")
        response.Write("</div></td>" & vbCRLF)
		response.write("</tr>" & vbCRLF)
'end
      response.write("</table></TD>" & vbCRLF)
  response.write("</TR>" & vbCRLF)
response.Write("</form>" & vbCRLF)
a = request("menu1")
   	if a="smd" or a="function_s" or a="push_s"or a="hMinia_s" or a="sensor_s" or a="" & _ 
		"rocker_s" or a="slide_s" or a="dip_s" or a="ear_j" or a="dc" or a="potentio" or a="" & _
		"connectors" or a="new"then
	show_result8 a
	elseif a ="a" then
  response.Write("<table width=100% border=0 cellspacing=0 cellpadding=0>"& vbCRLF) 
  response.Write("<TR>"& vbCRLF) 
  response.Write("<br>對不起,您未選擇搜尋的項目，<br>"& vbCRLF)
  response.Write("請選擇要搜尋的類別。<br><br>"& vbCRLF)
    response.Write("<TD width=492 height=18> <IMG SRC=images/data_14.gif WIDTH=492 HEIGHT=18></TD>"& vbCRLF)
    response.Write("<TD height=18 width=100% class=bg02>&nbsp; </TD>"& vbCRLF)
  response.Write("</TR>"& vbCRLF)
  response.Write("</table>"& vbCRLF) 
end if

b = request("menu81")
   	if b="smd" or b="function_s" or b="push_s"or b="hMinia_s" or b="sensor_s" or b="" & _ 
		"rocker_s" or b="slide_s" or b="dip_s" or b="ear_j" or b="dc" or b="potentio" or b="" & _
		"connectors" then
		show81 b
	elseif b ="a" then
  response.Write("<table width=100% border=0 cellspacing=0 cellpadding=0>"& vbCRLF) 
  response.Write("<TR>"& vbCRLF) 
  response.Write("<br>對不起,您未選擇搜尋的項目，<br>"& vbCRLF)
  response.Write("請選擇要搜尋的類別。<br><br>"& vbCRLF)
    response.Write("<TD width=492 height=18> <IMG SRC=images/data_14.gif WIDTH=492 HEIGHT=18></TD>"& vbCRLF)
    response.Write("<TD height=18 width=100% class=bg02>&nbsp; </TD>"& vbCRLF)
  response.Write("</TR>"& vbCRLF)
  response.Write("</table>"& vbCRLF) 
end if
	
c = request("search")
	if c <> "" then
	search c
	end if
end sub  

sub show81(b)
dim sql',show
set objConn = Server.CreateObject("Adodb.connection")
strCnn = "Driver={Microsoft Access Driver (*.mdb)};dbq=" & _
Server.MapPath("database/data_new.mdb")
objConn.Open strCnn
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "select search.name,search.link,search.id from search, main_table where main_table.table_list='"& b &"' and main_table.table_id=search.table"
	rs.Open sql,objConn,1,1
  response.Write("<table width=70% border=0 cellspacing=1 cellpadding=2>" & vbCRLF)	
	  response.Write("<tr>"& vbCRLF)
         response.Write("<td class=text04>您所搜尋的產品 <font color=#ff0000>"& b & "</font>，共有 ")
		 response.Write("<font color=#ff0000>" & rs.recordcount& "</font> 項相關資料。</td>" & vbCRLF)
	  response.Write("</tr>"& vbCRLF)
  response.write("<TR>" & vbCRLF) 
    response.write("<TD COLSPAN=2><br></td>" & vbCRLF)
        response.write("<tr>" & vbCRLF)
          response.write("<td width=40% class=text04>Name</td>" & vbCRLF)
          response.write("<td width=20% class=text04>Link</td>" & vbCRLF)
		response.write("</tr>" & vbCRLF)
	 response.write("</tr>" & vbCRLF)
	while Not rs.EOF
	response.Write("<tr>"& vbCRLF) 
	response.Write("<td width=40% class=text04><font color=#333333>" & rs("name") & "</font></td>"& vbCRLF)
  	response.Write("<td width=20% class=text04><font color=#333333>" & rs("link") & "</font></td>")
	response.Write("<td width=5% class=text04> <a href=d_eUpload.asp?id=" & rs("id") & "&name=search> <img src=images/d_change.gif width=36 height=24 border=0></a></td>"& vbCRLF)
	response.Write("<td width=5% class=text04> <a href=d_delete.asp?id=" & rs("id")& "&name=search> <img src=images/d_delete.gif width=36 height=24 border=0></a></td>"& vbCRLF)
	response.Write("</tr>"& vbCRLF)
	 	rs.MoveNext
		wend
	  response.Write("</table>" & vbCRLF)
     response.Write("<br>"& vbCRLF)
    response.Write("</TD>"& vbCRLF)
  response.Write("</TR>"& vbCRLF)
response.Write("<TABLE width=100% border=0 cellspacing=0 cellpadding=0>"& vbCRLF)
  response.Write("<TR>"& vbCRLF) 
    response.Write("<TD width=492 height=18> <IMG SRC=images/data_14.gif WIDTH=492 HEIGHT=18></TD>"& vbCRLF)
    response.Write("<TD height=18 class=bg02>&nbsp; </TD>"& vbCRLF)
  response.Write("</TR>"& vbCRLF)
response.Write("</TABLE>"& vbCRLF)
response.Write("</TABLE>"& vbCRLF)
rs.Close: Set rs = Nothing
objConn.close: Set objConn = Nothing
end sub

sub search(c)
dim strSql,n1',show,
strCnn = "Driver={Microsoft Access Driver (*.mdb)};dbq=" & _
    Server.MapPath("database/data_new.mdb")
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open strCnn
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "search", objConn, 1, 1
strSql = "name like '*" & c & "*'"				
rs.Filter = strSql
  response.write("<TR>" & vbCRLF) 
    response.write("<TD COLSPAN=2><br>" & vbCRLF)
      response.write("<table width=70% border=0 cellspacing=1 cellpadding=2>" & vbCRLF)
	  response.Write("<tr>"& vbCRLF)
         response.Write("<td class=text04>您所搜尋的產品 <font color=#ff0000>"& c & "</font>，共有 ")
		 response.Write("<font color=#ff0000>" & rs.recordcount& "</font>項相關資料。</td>" & vbCRLF)
	  response.Write("</tr>"& vbCRLF)
		response.write("<tr>" & vbCRLF)
	if rs.recordcount = "0" then 
		response.Write("<tr><td>"& vbCRLF) 
		response.Write("很抱歉，可能您輸入的型號有誤，<br>"& vbCRLF)
		response.Write("無法找到您所要的資料，您可重新進行查詢<br>"& vbCRLF)
		response.Write("</td></tr>"& vbCRLF) 
        response.write("<tr>" & vbCRLF)
	else
          response.write("<td width=40% class=text04>Name</td>" & vbCRLF)
          response.write("<td width=20% class=text04>Link</td>" & vbCRLF)
          response.write("<td width=5% class=text04>&nbsp;</td>" & vbCRLF)
          response.write("<td width=5% class=text04>&nbsp;</td>" & vbCRLF)
		response.write("</tr>" & vbCRLF)
     response.write("</table>" & vbCRLF)
  response.Write("<table width=70% border=0 cellspacing=1 cellpadding=2>" & vbCRLF)	
	while Not rs.EOF
	response.Write("<tr>"& vbCRLF) 
	response.Write("<td width=40% class=text04><font color=#333333>" & rs("name") & "</font></td>"& vbCRLF)
  	response.Write("<td width=20% class=text04><font color=#333333>" & rs("link") & "</font></td>")
	response.Write("<td width=5% class=text04> <a href=d_eUpload.asp?id=" & rs("id") & "&name=search> <img src=images/d_change.gif width=36 height=24 border=0></a></td>"& vbCRLF)
	response.Write("<td width=5% class=text04> <a href=d_delete.asp?id=" & rs("id")& "&name=search> <img src=images/d_delete.gif width=36 height=24 border=0></a></td>"& vbCRLF)
	response.Write("</tr>"& vbCRLF)
	 	rs.MoveNext
		wend
	end if
	  response.Write("</table>" & vbCRLF)
     response.Write("<br>"& vbCRLF)
    response.Write("</TD>"& vbCRLF)
  response.Write("</TR>"& vbCRLF)
  response.Write("<TR>"& vbCRLF) 
    response.Write("<TD width=492 height=18> <IMG SRC=images/data_14.gif WIDTH=492 HEIGHT=18></TD>"& vbCRLF)
    response.Write("<TD height=18 class=bg02>&nbsp; </TD>"& vbCRLF)
  response.Write("</TR>"& vbCRLF)
response.Write("</TABLE>"& vbCRLF)
rs.Close: Set rs = Nothing
objConn.close: Set objConn = Nothing

end sub

sub calculate()
dim a,b
a= request("choose")
b = request("record")
select case a
	case "8"
	response.Redirect("d_add8.asp")
	case "1"
		if b = "1" or b = "10" or b = "2" or b = "3" or b = "4" or b = "5" or b = "6" or b = "7" or b="" & _ 
		"8" or b = "9" then
			response.redirect("d_add1.asp?r="& b )
	end if
	response.Write("<div align=center>")
	'response.Write(b & "<br>")
	response.Write("你必需輸入號數而且是從1至10之間。謝謝合作！<br>")
	response.Write("<a href=javascript:history.back(-1);><img src=images/d_back.gif width=50 height=24 border=0></a>" & vbCRLF)
	response.Write("</div>")
end select
end sub

sub add_choose()
response.Write("<BODY BGCOLOR=#FFFFFF leftmargin=0 topmargin=0>" & vbCRLF)
response.Write("<Form name=form1 method=post action=manager.asp?do1=count1>"& vbCRLF)
response.Write("<TABLE WIDTH=100% BORDER=0 CELLPADDING=0 CELLSPACING=0>" & vbCRLF)
  response.Write("<TR> " & vbCRLF)
    response.Write("<TD width=492><img src=images/data_01.gif width=492 height=50 align=center></TD>" & vbCRLF)
    response.Write("<TD ROWSPAN=2 class=bg01>&nbsp; </TD>" & vbCRLF)
  response.Write("</TR>" & vbCRLF)
  response.Write("<TR> " & vbCRLF)
    response.Write("<TD width=492 height=37 background=images/batabg.gif>&nbsp;</TD>" & vbCRLF)
  response.Write("</TR>" & vbCRLF) 
  response.Write("<TR> " & vbCRLF)
    response.Write("<TD COLSPAN=2> <br>" & vbCRLF)
      response.Write("<table width=50% border=0 cellspacing=0 cellpadding=0>" & vbCRLF)
        response.Write("<tr> " & vbCRLF)
          response.Write("<td><div valign=bottom><img src=images/add.gif width=80 height=30></div></td>" & vbCRLF)
        response.Write("</tr>" & vbCRLF)
        response.Write("<tr>" & vbCRLF)
          response.Write("<td valign=center><table width=200 border=0 align=center cellpadding=2 cellspacing=0>" & vbCRLF)
              response.Write("<tr> " & vbCRLF)
                response.Write("<td colspan=3>&nbsp;" & vbCRLF)
              response.Write("</td></tr>" & vbCRLF)
              response.Write("<tr> " & vbCRLF)
                response.Write("<td width=20 valign=top> <input type=radio name=choose value=1> " & vbCRLF)
                response.Write("</td>" & vbCRLF)
                response.Write("<td width=39 class=text05>單頁</td>" & vbCRLF)
                response.Write("<td width=129 align=left><input name=record type=text size=2 maxlength=2>" & vbCRLF)
                  response.Write("<span class=text05> 筆</span></td>" & vbCRLF)
              response.Write("</tr>" & vbCRLF)
              response.Write("<tr> <td colspan=3><div align=right>" & vbCRLF)
			  response.Write("<a href=javascript:history.back(-1);><img src=images/d_back.gif width=50 height=24 border=0></a>" & vbCRLF)
                response.Write("<button name=submit type=submit width=36 height=24 border=0 class=input2><img src=images/d_submit.gif width=36 height=24>" & vbCRLF)
              response.Write("</button></div></td></tr>" & vbCRLF)
              response.Write("<tr> " & vbCRLF)
                response.Write("<td colspan=3>&nbsp;</td>" & vbCRLF)
              response.Write("</tr>" & vbCRLF)
            response.Write("</table>" & vbCRLF)
          response.Write("</td>" & vbCRLF)
        response.Write("</tr>" & vbCRLF)
      response.Write("</table> </TD>" & vbCRLF)
  response.Write("</TR>" & vbCRLF)
  response.Write("<TR> " & vbCRLF)
    response.Write("<TD width=492 height=18> <IMG SRC=images/data_14.gif WIDTH=492 HEIGHT=18></TD>" & vbCRLF)
    response.Write("<TD height=18 class=bg02>&nbsp; </TD>" & vbCRLF)
  response.Write("</TR>" & vbCRLF)
response.Write("</TABLE>" & vbCRLF)
response.Write("</FORM>" & vbCRLF)
end sub

sub show_result8(a)
set objConn = Server.CreateObject("Adodb.connection")
strCnn = "Driver={Microsoft Access Driver (*.mdb)};dbq=" & _
Server.MapPath("database/data_new.mdb")
objConn.Open strCnn
Set rs = Server.CreateObject("ADODB.Recordset")
  response.Write("<table width=70% border=0 cellspacing=1 cellpadding=2>" & vbCRLF)	
	rs.Open a,objConn,1,1
         response.Write("<td class=text04>您所搜尋的產品 <font color=#ff0000> "& a & "</font>，")
		 response.Write(" 共有 <font color=#ff0000>" & rs.recordcount& " </font>項相關資料。</td>" & vbCRLF)
	  response.Write("</tr>"& vbCRLF)
  response.write("<TR>" & vbCRLF) 
    response.write("<TD COLSPAN=2><br><td>" & vbCRLF)
	  response.Write("<tr>"& vbCRLF)
        response.write("<tr>" & vbCRLF)
          response.write("<td width=40% class=text04>Name</td>" & vbCRLF)
          response.write("<td width=20% class=text04>Link</td>" & vbCRLF)
		response.write("</tr>" & vbCRLF)
	while Not rs.EOF
	response.Write("<tr>"& vbCRLF) 
	response.Write("<td width=40% class=text04><font color=#333333>" & rs("name") & "</font></td>"& vbCRLF)
  	response.Write("<td width=20% class=text04><font color=#333333>" & rs("link") & "</font></td>")
	response.Write("<td width=5% class=text04> <a href=d_eUpload.asp?id=" & rs("id") & "&name="& a & "> <img src=images/d_change.gif width=36 height=24 border=0></a></td>"& vbCRLF)
	response.Write("<td width=5% class=text04> <a href=d_delete.asp?id=" & rs("id")& "&name="& a & "> <img src=images/d_delete.gif width=36 height=24 border=0></a></td>"& vbCRLF)
	response.Write("</tr>"& vbCRLF)
	 	rs.MoveNext
		wend
	  response.Write("</table>" & vbCRLF)
     response.Write("<br>"& vbCRLF)
    response.Write("</TD>"& vbCRLF)
  response.Write("</TR>"& vbCRLF)
response.Write("<TABLE width=100% border=0 cellspacing=0 cellpadding=0>"& vbCRLF)
  response.Write("<TR>"& vbCRLF) 
    response.Write("<TD width=492 height=18> <IMG SRC=images/data_14.gif WIDTH=492 HEIGHT=18></TD>"& vbCRLF)
    response.Write("<TD height=18 class=bg02>&nbsp; </TD>"& vbCRLF)
  response.Write("</TR>"& vbCRLF)
response.Write("</TABLE>"& vbCRLF)
response.Write("</TABLE>"& vbCRLF)
rs.Close: Set rs = Nothing
objConn.close: Set objConn = Nothing
end sub
%>
<HTML>
<HEAD>
<TITLE>database</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
<link href="hch.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.main {
	font-family: "新細明體";
	font-size: 12px;
	line-height: 18px;
	color: #999999;
}
-->
</style>
</HEAD>
<BODY BGCOLOR=#FFFFFF leftmargin="0" topmargin="0">
<p class="text01"> 
  <% main 
%>
</p>
</body>
</html>
