<!--#include file="Inc.asp"-->
<%
Dim Coin,Interest
If Not BBS.Founduser Then BBS.GoToErr(10)
Coin=BBS.Fun.GetStr("Coin")
If Coin<>"" Then If Not BBS.Fun.isInteger(Coin) then BBS.Alert"��������ȷ�Ľ�","?"
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
		'���»���
		BBS.UserLoginTrue()
		Rs.Close
		Set Rs=Nothing
End Sub


Sub Main
	BBS.Head"Bank.asp","","��̳����"
	Dim Content
	If Session(CacheName & "Bank")="" Then GetInterest
	Interest=Split(Session(CacheName & "Bank"),"|")
	Content="<table border=0 cellpadding=4 cellspacing=0 style='border-collapse: collapse' width='95%'><tr><td width='46%' align='center'><img src=images/bank.gif></td><td>"&_
	"<table align='center' width='95%' border='0' cellpadding='0' cellspacing='5' bgcolor='#FFFFFF' style='border-right: #BCBCBC 2px solid; border-bottom: #BCBCBC 2px solid;border-top: #e8e8e8 1px solid; border-left: #e8e8e8 1px solid;'>"&_
	"<tr><th height='25' bgcolor='#E8E8E8' colspan=2>����״��&nbsp;[�������:<font color=red>"&BBS.Info(119)*1000&"</font>��]</th></tr>"&_
	"<tr><td width='15%'>&nbsp;���л�����</td><td width='39%'><b>"&BBS.MyName&"</b></td></tr>"&_
	"<tr><td>&nbsp;���(����Ϣ)��</td><td><b><font color='red'>"&SESSION(CacheName & "MyInfo")(26)&"</font></b> Ԫ</td></tr>"&_
	"<tr><td>&nbsp;���ս�����Ϣ��</td><td><b><font color='red'>"&Interest(1)&"</font></b> Ԫ("&Interest(0)&"��)</td></tr>"&_
	"<tr><td>&nbsp;�����ֽ�</td><td><b><font color='Red'>"&SESSION(CacheName & "MyInfo")(7)&"</font></b> Ԫ</td></tr>"&_
	"<tr><td>&nbsp;�����ʽ��ܹ���</td><td><b><font color='red'>"&Ccur(SESSION(CacheName & "MyInfo")(26))+Ccur(SESSION(CacheName & "MyInfo")(7))&"</font></b> Ԫ</td></tr>"&_
	"</table></td></tr></table>"&_
	"<table width='95%' border='0' cellPadding='0' cellSpacing='0'><tr><td>"&_
	"<form action='?Action=Save' method='post'><table align='center' width='95%' border='0' cellpadding='0' cellspacing='5' bgcolor='#FFFFFF' style='border-right: #BCBCBC 2px solid; border-bottom: #BCBCBC 2px solid;border-top: #e8e8e8 1px solid; border-left: #e8e8e8 1px solid;'>"&_
	"<tr><th height='25' bgcolor='#E8E8E8'>����̨</th></tr>"&_
	"<tr><td>&nbsp;�ֽ�<b><font color=red>"&SESSION(CacheName & "MyInfo")(7)&"</font></b> Ԫ</td></tr>"&_
	"<tr><td align='center'>�洢��<input size='10' value='1000' name='Coin'>Ԫ&nbsp;&nbsp;<input type='submit'  value=' ��Ǯ '></td></tr></table></form>"&_
	"</td><td>"&_
	"<form action='?Action=Draw' method='post'><table align='center' width='95%' border='0' cellpadding='0' cellspacing='5' bgcolor='#FFFFFF' style='border-right: #BCBCBC 2px solid; border-bottom: #BCBCBC 2px solid;border-top: #e8e8e8 1px solid; border-left: #e8e8e8 1px solid;'>"&_
	"<tr><th height='25' bgcolor='#E8E8E8'>֧ȡ��̨</th></tr>"&_
	"<tr><td>&nbsp;��<b><font color=Red>"&SESSION(CacheName & "MyInfo")(26)&"</font></b> Ԫ</td></tr>"&_
	"<tr><td align='center'>֧ȡ <input size='10' value='1000' name='Coin'> Ԫ&nbsp;&nbsp;<input type='submit'  value=' ֧ȡ '></td></tr></table></form>"&_
	"</td><td>"&_
	"<form action='?Action=Virement' method='post'><table align='center' width='95%' border='0' cellpadding='0' cellspacing='5' bgcolor='#FFFFFF' style='border-right: #BCBCBC 2px solid; border-bottom: #BCBCBC 2px solid;border-top: #e8e8e8 1px solid; border-left: #e8e8e8 1px solid;'>"&_
	"<tr><th height='25' bgcolor='#E8E8E8'>ת�ʹ�̨</th></tr>"&_
	"<tr><td>&nbsp;�Ѵ��ת�ʸ�����</td></tr>"&_
	"<tr><td align='center'><input size='5' value='1000' name='Coin'>Ԫ&nbsp;��<input size='5' name='ToUserName'>&nbsp<input type='submit' value=' ת�� '></td></tr></table></form>"&_
	"</td></tr></table>"
	Call BBS.ShowTable("��̳����",Content)
End Sub

Sub Save
	If Coin="" Then BBS.Alert"ʧ�ܣ�����û����дҪ���Ľ�","?"
	If Int(Coin) > Int(SESSION(CacheName & "MyInfo")(7)) Then BBS.Alert "ʧ�ܣ��������ȿ�����ڴ����ж���Ǯ�в��У����ϰ�һëǮ��һ�٣�","?"
	BBS.Execute("update [user] Set BankSave=BankSave+"&Coin&",Coin=Coin-"&Coin&" where Name='"&BBS.MyName&"'")
	Session(CacheName & "MyInfo") = Empty
	BBS.Alert"��ϲ�����д��ɹ�","?"
End sub

Sub Draw
	If Coin="" Then BBS.Alert"ʧ�ܣ�����û����дҪȡ��Ľ�","?"
	If int(Coin) > Int(SESSION(CacheName & "MyInfo")(26)) Then BBS.Alert "Ү����λƯ����������Ӣ����С�㣬������Ӣ���ޱȣ��������кܺ�����С�ķ�ë��ѽ�㣡����ע����Ĵ��ж��٣�","?"
	BBS.Execute("update [user] Set BankSave=BankSave-"&Coin&",Coin=Coin+"&Coin&" where Name='"&BBS.MyName&"'")
	Session(CacheName & "MyInfo") = Empty
	BBS.Alert"��ϲ������ȡ��ɹ���","?"
End Sub

Sub Virement
	Dim ToUserName,Sms,TmpUbbString
	ToUserName=BBS.Fun.GetStr("ToUserName")
	IF Not BBS.Fun.CheckIn(ToUserName) Or ToUserName="" Then BBS.Alert"ʧ�ܣ����޴���,�����ʵ�����ͳ�Ǯ���Ļ�,�͸�վ���ͺ��ˣ�","?"
	If Int(Coin) > Int(SESSION(CacheName& "MyInfo")(26)) then BBS.Alert "ʧ�ܣ����ǲ��Ǻ��Ĺ�ͷ���������˼Ҵ�ծ�ߵ�̫���������Լ����ж���Ǯ����֪������","?"
	If LCase(ToUserName)=LCase(BBS.MyName) Then BBS.Alert "ʧ�ܣ����Լ�ת�ʺܺ�����","?"
	IF BBS.Execute("Select Name From[user] where Name='"&ToUserName&"'").Eof Then
 	  BBS.Alert"ʧ�ܣ����޴���,�����ʵ�����ͳ�Ǯ���Ļ�,�͸�վ���ͺ��ˣ�","?"
	End if
	BBS.Execute("Update [user] Set BankSave=BankSave-"&Coin&" where Name='"&BBS.MyName&"'")
	Sms="���µ��´��ڱ���,"&BBS.MyName&"ͨ������ת��������"&Coin&"Ԫ�ֽ������Ե��������й�̨���գ�"&vbcrlf&"<div align='right' style='color:#F00'>���������С��Զ�����ϵͳ</div>"
	BBS.Execute("Update [user] Set BankSave=BankSave+"&Coin&",NewSmsNum=NewSmsNum+1,SmsSize=SmsSize+1 where Name='"&ToUserName&"'")
	BBS.execute("insert Into [Sms](Name,Content,MyName,MyFlag)VALUES('�Զ�����ϵͳ','"&Replace(Sms,"'","''")&"','"&ToUserName&"',1)")
	BBS.UpdageOnline ToUserName,1
	Session(CacheName & "MyInfo") = Empty
	BBS.Alert"��ϲ��ת�ʳɹ�,ϵͳ���Զ�����֪ͨ���������ѣ�","?"
End sub

%>