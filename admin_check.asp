<!--#include file="Inc.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="inc/Style.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/javascript">
function checkclick(msg,url){if(confirm(msg))window.location.href=url;}
function openwin(url,w,h,s){window.open(url,'_blank','status=yes,scrollbars='+s+',top=20,left=110,width='+w+',height='+h);}
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name == 'ID'){
	e.checked = form.chkall.checked;
	}
   }
  }
</script>
<%
Server.ScriptTimeOut=99999
Const IconA="<img border=""0"" align=""absmiddle"" alt=""添加"" src=""Images/icon/add.gif"" /> "
Const IconE="<img border=""0"" align=""absmiddle"" alt=""编辑"" src=""Images/icon/edit.gif"" /> "
Const IconD="<img border=""0"" align=""absmiddle"" alt=""删除"" src=""Images/icon/del.gif"" /> "
Const IconH="<img border=""0"" align=""absmiddle"" alt=""帮助"" src=""Images/icon/help.gif"" /> "
Dim Rs,AdminString
CheckAdmin()

Sub CheckAdmin()
	Dim AdminName,AdminPassword
	AdminName=BBS.GetMemor("Admin","AdminName")
	AdminPassword=BBS.GetMemor("Admin","AdminPassword")
	IF AdminName="" or AdminPassword="" Then
		If Instr(PageURL,"admin_index.asp")>0 Then
			Response.redirect"admin_login.asp"
		Else
		Head
		ShowTable"限制进入","您还不是管理员，请先 【<a href='Admin_Login.asp' target='_parent'>登陆</a>】。"
		BBS.NetLog"！非法尝试进入后台失败!"
		Footer
		Response.end
		End If
	End If
	If not BBS.Fun.CheckName(AdminName) or not BBS.Fun.CheckPassword(AdminPassword) then
		Response.redirect"Admin_login.asp"
		Response.end
	End If
	If BBS.Execute("Select name from [Admin] where Name='"&AdminName&"' and Password='"&AdminPassword&"' and boardID=0 ").eof  Then
		Response.redirect"admin_login.asp"
		Response.end
	End if
	BBS.MyName=AdminName
	AdminString=BBS.execute("select Strings From [Admin] where name='"&AdminName&"' And Password='"&AdminPassword&"' And boardID=0")(0)
End Sub

Sub CheckString(Flag)
	If Instr(AdminString,","&Flag&",")=0 Then
		Goback"对不起","你没有该页的操作权限!"
		Footer	
		Response.end
	End If
End Sub
Sub GoBack(Str1,Str2)
	If Str1="" Then Str1="错误信息"
	If Str2="" Then Str2="请填写完整必填项目"
	Str2=Str2&" <a href=""javascript:history.go(-1)"">返回重填</a></li>"
	Response.Write"<div class=""mian""><div class=""top"">"&Str1&" </div><div class=""divtr1"" style=""height:50px; ""><div class=""divtd1"" style=""color:red;FONT: 50px/50px 宋体;height:50px;"">×</div><div class=""divtd1"" style=""margin-top:8px;"">"&str2&"</div></div></div>"
End Sub

Function GetRadio(Input_name,txt_Name,A,B)
	Dim temp
On Error Resume Next'调试
	If A="" Then A=0
	If Int(A)=Int(B) then temp="checked "
	GetRadio=" <input type='radio' name='"&Input_name&"' value='"&B&"' "&Temp&"/>"&txt_name&""
'if err then Response.Write Input_name:Response.end
End function

Sub ShowTable(Str1,Str2)
	Response.Write"<div class='mian'><div class='top'>"&Str1&" </div><div class='divtr1' style='padding:10px;line-height: 24px'>"&str2&"</div></div>"
End Sub

Sub Suc(Str1,Str2,url)
	If Str1="" Then Str1="操作成功"
	If Str2="" Then Str2="成功的完成这次操作！"
	Str2=Str2&"<a href="""&Url&""" >返回继续管理</a>"
	Response.Write"<div class=""mian""><div class=""top"">"&Str1&" </div><div class=""divtr1"" style=""height:50px; ""><div class=""divtd1"" style=""color:red;FONT: 50px/50px 宋体;height:50px;"">√</div><div class=""divtd1"" style=""margin-top:8px;"">"&str2&"</div></div></div>"
End Sub

Sub Head()
	Response.Write"</head><body>"
End Sub

Sub Footer()
	Response.Write"</body></html>"
	Set Rs=Nothing
	Set BBS =Nothing
End Sub

Sub DIVTR(T1,T2,Str,H,show)
	Dim StyleH
	If T2<>"" Then T2="<div>"&T2&"</div>"
	StyleH="min-height:"&H&"px;"
	'识别IE浏览器
	If BBS.MSIE Then StyleH=Replace(StyleH,"min-","")
	Response.Write"<div class='divtr"&Show&"'><div style='width:200px;"&StyleH&"float:left;'><div class='title'>"&T1&"</div>"&T2&"</div><div style='text-align :left"&StyleH&"'><div>"&str&"</div></div><div style='clear: both;'></div></div>"
End Sub
%>
