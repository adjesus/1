<!--#include file="Inc.asp"-->
<!--#include file="inc/md5.asp"-->
<%Dim Action
Action=request.querystring("Action")
Select Case Action
Case"exit"
	ExitLogin()
Case"login"
	CheckLogin()
Case Else
	Main()
End select
Set BBS =Nothing

Sub Main()
	Response.Write"<link rel=stylesheet type=text/css href='Inc/Style.css' />"&_
	"<title>�����½</title>"&_
    "<form method=POST action='?action=login'>"&_
	"<div class='mian1' style='width:400px'>"&_
	"<div class='top'>�����½</div>"&_
	"<div class='divtr1' style='height:100px'>"&_
	"<div style='margin-left:50px'>�û����ƣ�<input name='AdminName' type=text size='22' style='width:120px'></div>"&_
	"<div style='margin-left:50px'>��̨���룺<input name='Password' type=password size='22' style='width:120px'></div>"&_
	"<div style='margin-left:50px'>��֤���룺"&BBS.GetiCode&"</div> "&_
	"<div class='bottom'><input type='submit' class='button' value='�� ¼'></div></div></form>"
End Sub

Sub CheckLogin()
	Dim AdminName,AdminPassword,PassCode,Temp
	With BBS
	AdminName=.Fun.GetStr("AdminName")
	AdminPassword=.Fun.GetStr("Password")
	PassCode=.Fun.GetStr("iCode")
	If PassCode="" or AdminName="" or AdminPassword="" Then .Alert"���������������ύ��","admin_login.asp"
	If not .Fun.CheckName(AdminName) or not .Fun.CheckPassword(AdminPassword) then
		.SafeBuckler AdminName,.MyIP,1
		.Alert"��������û��������ڻ����������","admin_login.asp"
	End if
	AdminPassword=Md5(AdminPassword)
	If .SafeBuckler(AdminName,.MyIP,0) Then .Alert"BBS��ȫ����������\n�Բ����㳢�Ե�½���󳬹�5�Σ����첻���ٵ�½��̨��\n�����Ϣ�ѱ�ϵͳ��¼��","Index.asp"
	If .execute("select name From [Admin] where name='"&AdminName&"' And Password='"&AdminPassword&"' And boardID=0").eof Or Session("iCode")<>PassCode  Then
		Session("iCode")=Empty
		.SafeBuckler AdminName,.MyIP,1
		.Alert"��������û��������ڻ������������������֤�����","Admin_login.asp"
	Else
		.LetMemor "Admin","AdminName",AdminName
		.LetMemor "Admin","AdminPassword",AdminPassword
		.MyName=AdminName
		.NetLog"�ɹ���½��̨"
		If .Info(16)="1" Then .execute("delete from [Log] where DATEDIFF('d', LogTime,'"&.NowBBSTime&"')>7")
		Session("iCode")=Empty
		Response.Redirect"admin_index.asp"
		Response.End
	End if
	End With
End Sub

Sub ExitLogin()
	Session(CacheName &"AdminName") = Empty
	Response.Cookies(CacheName &"Admin")("AdminName")= Empty
	Session(CacheName &"AdminName") = Empty
	Response.Cookies(CacheName &"Admin")("AdminPassword")= Empty
	'Session.Abandon
	Response.redirect"Index.asp"
End Sub
%>