<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
' ==========================================================
' HCH Database Setup Script
' Run this ONCE to create the messages table in data_new.mdb
' Access via: http://yoursite/db_setup.asp?key=hchsetup2024
' ==========================================================

Dim secretKey : secretKey = "hchsetup2024"
If Request.QueryString("key") <> secretKey Then
  Response.Write "<h3 style='color:red;font-family:Arial;'>Access Denied. Provide valid key.</h3>"
  Response.End
End If

Response.ContentType = "text/html"
Response.Charset = "UTF-8"
%>
<!DOCTYPE html>
<html><head><meta charset="UTF-8"><title>HCH DB Setup</title>
<style>body{font-family:Arial;max-width:700px;margin:40px auto;padding:20px;}
.ok{color:green;}.err{color:red;}.info{color:#333;}
pre{background:#f5f5f5;padding:12px;border-radius:4px;font-size:13px;}</style>
</head><body>
<h2>HCH Database Setup</h2>
<%
Dim conn, sql, result
Set conn = Server.CreateObject("ADODB.Connection")
On Error Resume Next
conn.Open "Driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath("database/data_new.mdb")
If Err.Number <> 0 Then
  Response.Write "<p class='err'>❌ Cannot open database: " & Err.Description & "</p>"
  Response.End
End If
Response.Write "<p class='ok'>✅ Database connected successfully.</p>"

' Create messages table
Err.Clear
sql = "CREATE TABLE messages (" & _
      "id AUTOINCREMENT PRIMARY KEY," & _
      "msg_name TEXT(100)," & _
      "msg_email TEXT(200)," & _
      "msg_phone TEXT(50)," & _
      "msg_company TEXT(200)," & _
      "msg_partno TEXT(200)," & _
      "msg_subject TEXT(200)," & _
      "msg_content MEMO," & _
      "msg_read BYTE DEFAULT 0," & _
      "msg_date DATETIME" & _
      ")"
conn.Execute(sql)
If Err.Number = 0 Then
  Response.Write "<p class='ok'>✅ Table [messages] created successfully.</p>"
ElseIf InStr(Err.Description, "already exists") > 0 Or Err.Number = -2147467259 Then
  Response.Write "<p class='info'>ℹ️ Table [messages] already exists — skipped.</p>"
Else
  Response.Write "<p class='err'>⚠️ Table [messages]: " & Err.Description & " (Error " & Err.Number & ")</p>"
End If

conn.Close: Set conn = Nothing
On Error GoTo 0
%>
<hr>
<h3>Setup Complete</h3>
<p>You may now:</p>
<ul>
  <li><a href="index.asp">Visit the homepage</a></li>
  <li><a href="contact.asp">Test the contact form</a></li>
  <li><a href="admin/messages.asp">View the admin panel</a> (login: hchlogin / looklogin)</li>
</ul>
<p style="color:#999;font-size:12px;margin-top:20px;">⚠️ Delete or rename this file (db_setup.asp) after setup is complete for security.</p>
</body></html>
