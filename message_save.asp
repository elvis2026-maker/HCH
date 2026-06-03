<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Response.ContentType = "text/plain"
Response.Charset = "UTF-8"

Dim fName, fEmail, fPartno, fMsg
fName   = Trim(Request.Form("chat_name"))
fEmail  = Trim(Request.Form("chat_email"))
fPartno = Trim(Request.Form("chat_partno"))
fMsg    = Trim(Request.Form("chat_msg"))

If fName = "" Or fEmail = "" Or fMsg = "" Then
  Response.Write "ERROR:MISSING"
  Response.End
End If

On Error Resume Next
Dim conn
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath("database/data_new.mdb")

If Err.Number <> 0 Then
  Response.Write "ERROR:DB"
  Response.End
End If

Dim sql
sql = "INSERT INTO messages (msg_name, msg_email, msg_phone, msg_company, msg_partno, msg_subject, msg_content, msg_date) VALUES ('" & _
  Replace(fName,"'","''") & "','" & _
  Replace(fEmail,"'","''") & "','','','" & _
  Replace(fPartno,"'","''") & "','LINE Widget','" & _
  Replace(fMsg,"'","''") & "',Now())"

conn.Execute(sql)

If Err.Number <> 0 Then
  conn.Close: Set conn = Nothing
  Response.Write "ERROR:INSERT"
  Response.End
End If

conn.Close: Set conn = Nothing
On Error GoTo 0

Response.Write "OK"
%>
