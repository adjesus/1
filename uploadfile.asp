<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style>
BODY {FONT: 12px/25px sans-serif;margin:0}
form{margin:0px}
input{border:#999999 1px solid;
background:#FFFFFF;
height:20px;
margin:0 5px 0 0}
</style>
</head>
<%
Dim Flag
If Request("Flag") = "" Then  Flag= 0 Else Flag = 1
Response.Write"<body><form action='SaveUploadFile.asp?Flag="&Flag&"' method='post' enctype='multipart/form-data' target='_self'><input name='filedata' type='file' size='18'><input type='Submit' value='上传' name='Submit' id='Submit'></form>"
%>
</body>
</html>