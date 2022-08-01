<!--#include file="Inc.asp"-->
<!--#include file="Inc/Md5.asp"-->
<%
BBS.Head"login.asp","","论坛登陆"
Dim Action
Action=Lcase(Request.querystring("action"))
If len(Action)>10 Then BBS.GotoErr(1)
Select Case Action
	Case"exit"
		If Request.ServerVariables("request_method") <> "POST" then
		Response.write "<form name='KK' method=post action=?"&Request.ServerVariables("Query_String")&"></form><SCRIPT>if(confirm('您确定要退出论坛么?')){returnValue=KK.submit()}else{returnValue=history.back()}</SCRIPT>"
		Response.End
		End If
	ExitLogin()
	Case"login":CheckLogin()
	Case else:Mian()
End select
BBS.Footer()
Set BBS =Nothing


Sub Mian()
Dim S
S=Request.ServerVariables("HTTP_REFERER")
If instr(lcase(S),"login.asp")>0 or instr(lcase(S),"err.asp")>0 then
Else
Session(CacheName&"BackURL")=S
End If
S="<form method=""post"" style=""margin:0px"" action=""login.asp?action=login"">"
S=S&BBS.Row("<b>请输入您的用户名：</b>","<input name=""name"" type=""text"" class=""submit"" size=""20"" /> <a href=""register.asp"">没有注册？</a>","65%","")
S=S&BBS.Row("<b>请输入您的密码：</b>","<input name=""Password"" type=""password"" size=""20"" /> <a href=""usersetup.asp?action=forgetpassword"">忘记密码？</a>","65%","")
If BBS.Info(14)="1" Then
	S=S&BBS.Row("<b>请输入右边的验证码：</b>",BBS.GetiCode,"65%","")
Else
	S=S&"<input name=""iCode"" type=""hidden"" value=""BBS"" />"
End If
S=S&BBS.Row("<b>Cookie 选项：</b>","<input type=radio  name=""cookies"" value=""0"" checked class=checkbox />不保存 <input type=radio  name=cookies value=""1"" class=checkbox />保存一天 <input type=radio  name=cookies value=""30"" class=checkbox />保存一月","65%","")
S=S&BBS.Row("<b>选择登陆方式：</b>","<input type=radio value=""1"" checked name='hidden' class=checkbox />正常登陆 <input type='radio' value='2' name='hidden' class=checkbox />隐身登陆","65%","")
S=S&"<div style="" padding:5px;BACKGROUND: "&BBS.SkinsPIC(2)&";"" align=""center""><input Class='login' type=""submit"" value=""登 陆"" /></div></form>"
BBS.ShowTable"用户登陆",S
End Sub

Sub CheckLogin()
	With BBS
	Dim Rs,UserName,Password,IsHidden,CookiesDate,Content,iCode,S
	.CheckMake
	If .Info(10)<>"0" Then
		If Session(CacheName&"LoginTime")+.Info(10)/1440>now() then .GotoErr(5)
	End If
	UserName=Request.Form("name")
	Password=Request.Form("password")
	IsHidden=Request.Form("hidden")
	iCode=Request.Form("iCode")
	CookiesDate=.CheckNum(Request.Form("cookies"))
	If UserName="" or Password="" Then .GoToErr(6)
	If .Info(14)="1" Then
		If iCode<>Session("iCode") or Session("iCode")="" Then .GotoErr(8)
	End If
	If Not .Fun.Checkname(UserName) OR Not .Fun.CheckPassword(Password) then .GotoErr(16)
	If .SafeBuckler(UserName,.MyIP,0) Then .Alert"BBS安全盾已启动！\n\n对不起，你尝试登陆错误超过3次，今天不能再登陆论坛。\n你的信息已被系统记录！","Index.asp"
	Password=MD5(Password)
	Set Rs = .Execute("select top 1 Id,Name,Password,Isdel,LastTime From [User] where name='"&UserName&"' and password='"&Password&"' and (Isdel=0 or Isdel=2)")
	If Rs.Eof then
		.SafeBuckler UserName,.MyIP,1
		.GotoErr(9)
	ElseIf Rs(3)=2 Then
		.GotoErr(78)
	Else
		.LetMemor "","MyID",Rs(0)
		.LetMemor "","MyName",Rs(1)
		.LetMemor "","MyPassword",Password
		.LetMemor "","MyHidden",IsHidden
		.LetMemor "","CookiesDate",CookiesDate
		.LetMemor "","LastTime",Rs(4)
		Session(CacheName & "login")="1"
		If Int(CookiesDate)>0 Then	Response.Cookies(CacheName).Expires=date+Int(CookiesDate)
		.Execute("update [user] set LastTime='"&.NowBbsTime&"',LastIp='"&.MyIp&"' where ID="&Rs(0))
		Session(CacheName&"LoginTime")=now()
		S=Session(CacheName&"BackURL")
		If S="" Then S="Index.asp"
		Content="<meta http-equiv=refresh content=2;url=Index.asp><div style='margin: 15px;line-height: 150%'><li><b>3</b> 秒钟后将自动返回首页</li><li><a href='Index.asp'>立即进入论坛首页</a></li><li><a href="&S&">返回刚才浏览的页面</a><br></div>"
	End if
	Rs.Close
	Set Rs=Nothing
	.ShowTable"登陆成功",Content
	End With
End Sub

Sub ExitLogin()
	BBS.SetMemorEmpty()
	BBS.ShowTable "退出论坛","<div style='margin: 15px;line-height: 150%'><li>已经成功的退出论坛</li><li><a href='login.asp'>重新登陆</a></li><li><a href='Index.asp'>进入论坛首页</a></li></div>"
End Sub
%>