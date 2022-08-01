<!--#include file="Inc.asp"-->
<%
Dim Coin,Interest
If Not BBS.Founduser Then BBS.GoToErr(10)
Coin=BBS.Fun.GetStr("Coin")
If Coin<>"" Then If Not BBS.Fun.isInteger(Coin) then BBS.Alert"请输入正确的金额！","?"
Select Case Request.querystring("Action")
Case "Save"
	Save
Case "Draw"
	Draw
Case "Virement"
	Virement
Case"Convert"
	Convert
Case Else
	Main
End Select
BBS.Footer()
Set BBS =Nothing

Sub GetInterest()
	Dim Rs,Temp
		Set Rs=BBS.Execute("select Coin,BankSave,Banktime from [user] where ID="&BBS.MyID&"")
		If Rs.Eof Then
			BBS.SetMemorEmpty()
			BBS.GoToErr(4)
		End If
		Temp=ccur(ccur(rs(1))*ccur(Datediff("d",cdate(rs("banktime")),cdate(BBS.NowBbsTime)))*BBS.Info(119))
		Session(CacheName & "Bank")=Datediff("d",cdate(rs(2)),cdate(BBS.NowBbsTime))&"|"&Temp
		BBS.Execute("Update [user] Set BankSave=BankSave+"&Temp&",BankTime='"&BBS.NowBbsTime&"' where Name='"&BBS.MyName&"' ")
		Session(CacheName & "MyInfo") = Empty
		'更新缓存
		BBS.UserLoginTrue()
		Rs.Close
		Set Rs=Nothing
End Sub


Sub Main
	BBS.Head"Bank.asp","","论坛银行"
	Dim Content
	If Session(CacheName & "Bank")="" Then GetInterest
	Interest=Split(Session(CacheName & "Bank"),"|")
	Content="<table border=0 cellpadding=4 cellspacing=0 style='border-collapse: collapse' width='95%'><tr><td width='46%' align='center'><img src=images/bank.gif></td><td>"&_
	"<table align='center' width='95%' border='0' cellpadding='0' cellspacing='5' bgcolor='#FFFFFF' style='border-right: #BCBCBC 2px solid; border-bottom: #BCBCBC 2px solid;border-top: #e8e8e8 1px solid; border-left: #e8e8e8 1px solid;'>"&_
	"<tr><th height='25' bgcolor='#E8E8E8' colspan=2>财务状况&nbsp;[存款利率:<font color=red>"&BBS.Info(119)*1000&"</font>‰]</th></tr>"&_
	"<tr><td width='15%'>&nbsp;银行户主：</td><td width='39%'><b>"&BBS.MyName&"</b></td></tr>"&_
	"<tr><td>&nbsp;存款(含利息)：</td><td><b><font color='red'>"&SESSION(CacheName & "MyInfo")(26)&"</font></b> 元</td></tr>"&_
	"<tr><td>&nbsp;今日结算利息：</td><td><b><font color='red'>"&Interest(1)&"</font></b> 元("&Interest(0)&"天)</td></tr>"&_
	"<tr><td>&nbsp;持有现金：</td><td><b><font color='Red'>"&SESSION(CacheName & "MyInfo")(7)&"</font></b> 元</td></tr>"&_
	"<tr><td>&nbsp;个人资金总共：</td><td><b><font color='red'>"&Ccur(SESSION(CacheName & "MyInfo")(26))+Ccur(SESSION(CacheName & "MyInfo")(7))&"</font></b> 元</td></tr>"&_
	"</table></td></tr></table>"&_
	"<table width='95%' border='0' cellPadding='0' cellSpacing='0'><tr><td>"&_
	"<form action='?Action=Save' method='post'><table align='center' width='95%' border='0' cellpadding='0' cellspacing='5' bgcolor='#FFFFFF' style='border-right: #BCBCBC 2px solid; border-bottom: #BCBCBC 2px solid;border-top: #e8e8e8 1px solid; border-left: #e8e8e8 1px solid;'>"&_
	"<tr><th height='25' bgcolor='#E8E8E8'>存款柜台</th></tr>"&_
	"<tr><td>&nbsp;现金：<b><font color=red>"&SESSION(CacheName & "MyInfo")(7)&"</font></b> 元</td></tr>"&_
	"<tr><td align='center'>存储：<input size='10' value='1000' name='Coin'>元&nbsp;&nbsp;<input type='submit'  value=' 存钱 '></td></tr></table></form>"&_
	"</td><td>"&_
	"<form action='?Action=Draw' method='post'><table align='center' width='95%' border='0' cellpadding='0' cellspacing='5' bgcolor='#FFFFFF' style='border-right: #BCBCBC 2px solid; border-bottom: #BCBCBC 2px solid;border-top: #e8e8e8 1px solid; border-left: #e8e8e8 1px solid;'>"&_
	"<tr><th height='25' bgcolor='#E8E8E8'>支取柜台</th></tr>"&_
	"<tr><td>&nbsp;存款：<b><font color=Red>"&SESSION(CacheName & "MyInfo")(26)&"</font></b> 元</td></tr>"&_
	"<tr><td align='center'>支取 <input size='10' value='1000' name='Coin'> 元&nbsp;&nbsp;<input type='submit'  value=' 支取 '></td></tr></table></form>"&_
	"</td><td>"&_
	"<form action='?Action=Virement' method='post'><table align='center' width='95%' border='0' cellpadding='0' cellspacing='5' bgcolor='#FFFFFF' style='border-right: #BCBCBC 2px solid; border-bottom: #BCBCBC 2px solid;border-top: #e8e8e8 1px solid; border-left: #e8e8e8 1px solid;'>"&_
	"<tr><th height='25' bgcolor='#E8E8E8'>转帐柜台</th></tr>"&_
	"<tr><td>&nbsp;把存款转帐给好友</td></tr>"&_
	"<tr><td align='center'><input size='5' value='1000' name='Coin'>元&nbsp;给<input size='5' name='ToUserName'>&nbsp<input type='submit' value=' 转帐 '></td></tr></table></form>"&_
	"</td></tr></table>"
	Call BBS.ShowTable("论坛银行",Content)
End Sub

Sub Save
	If Coin="" Then BBS.Alert"失败！您还没有填写要存款的金额！","?"
	If Int(Coin) > Int(SESSION(CacheName & "MyInfo")(7)) Then BBS.Alert "失败！拜托你先看看你口袋里有多少钱行不行？别老把一毛钱当一百！","?"
	BBS.Execute("update [user] Set BankSave=BankSave+"&Coin&",Coin=Coin-"&Coin&" where Name='"&BBS.MyName&"'")
	Session(CacheName & "MyInfo") = Empty
	BBS.Alert"恭喜！银行存款成功","?"
End sub

Sub Draw
	If Coin="" Then BBS.Alert"失败！您还没有填写要取款的金额！","?"
	If int(Coin) > Int(SESSION(CacheName & "MyInfo")(26)) Then BBS.Alert "耶！这位漂亮的先生或英俊的小姐，您真是英勇无比，请问银行很好抢吗？小心飞毛腿呀你！（请注意你的存款还有多少）","?"
	BBS.Execute("update [user] Set BankSave=BankSave-"&Coin&",Coin=Coin+"&Coin&" where Name='"&BBS.MyName&"'")
	Session(CacheName & "MyInfo") = Empty
	BBS.Alert"恭喜！银行取款成功！","?"
End Sub

Sub Virement
	Dim ToUserName,Sms,TmpUbbString
	ToUserName=BBS.Fun.GetStr("ToUserName")
	IF Not BBS.Fun.CheckIn(ToUserName) Or ToUserName="" Then BBS.Alert"失败！查无此人,如果您实在想送出钱来的话,送给站长就好了！","?"
	If Int(Coin) > Int(SESSION(CacheName& "MyInfo")(26)) then BBS.Alert "失败！你是不是好心过头啦？还是人家催债催得太紧，让你自己还有多少钱都不知道？！","?"
	If LCase(ToUserName)=LCase(BBS.MyName) Then BBS.Alert "失败！给自己转帐很好玩吗？","?"
	IF BBS.Execute("Select Name From[user] where Name='"&ToUserName&"'").Eof Then
 	  BBS.Alert"失败！查无此人,如果您实在想送出钱来的话,送给站长就好了！","?"
	End if
	BBS.Execute("Update [user] Set BankSave=BankSave-"&Coin&" where Name='"&BBS.MyName&"'")
	Sms="天下掉下大馅饼啦,"&BBS.MyName&"通过友情转帐赠送您"&Coin&"元现金！您可以到社区银行柜台查收！"&vbcrlf&"<div align='right' style='color:#F00'>「社区银行」自动送信系统</div>"
	BBS.Execute("Update [user] Set BankSave=BankSave+"&Coin&",NewSmsNum=NewSmsNum+1,SmsSize=SmsSize+1 where Name='"&ToUserName&"'")
	BBS.execute("insert Into [Sms](Name,Content,MyName,MyFlag)VALUES('自动送信系统','"&Replace(Sms,"'","''")&"','"&ToUserName&"',1)")
	BBS.UpdageOnline ToUserName,1
	Session(CacheName & "MyInfo") = Empty
	BBS.Alert"恭喜！转帐成功,系统已自动发信通知了您的朋友！","?"
End sub

%>