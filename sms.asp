<!--#include file="Inc.asp"-->
<!--#include file="Inc/Page_Cls.asp"-->
<!--#include file="inc/ubb_Cls.asp"-->
<%
Dim action,AllSmsSize
If Not BBS.FoundUser Then BBS.GoToErr(20)
BBS.Position=BBS.Position&" -> <a href=""userinfo.asp"">�û��������</a>"
action=Lcase(Request("action"))
BBS.Head "sms.asp","","�����ż�"
ShowMySmsInfo()
If Len(action)>10 Then BBS.GoToErr(1)
Select Case action
Case"save"
	SaveSms
Case"del"
	Del
Case"delall"
	DelAll
Case"write"
	WriteSms
Case Else
	ReadSms
End Select
BBS.Footer()
Set BBS =Nothing


Sub Del
	Dim ID,I,Rs
	ID=BBS.CheckNum(request("ID"))
	Set Rs=BBS.Execute("Select MyName,Name From[sms] where ID="&ID&" And (Name='"&BBS.MyName&"' or MyName='"&BBS.MyName&"')")
	If not Rs.eof then
		If Lcase(BBS.MyName)=Lcase(Rs(0)) Then
			BBS.execute("Update [sms] set MyFlag=2 where ID="&ID)
		Else
			BBS.execute("Update [sms] set Flag=2 where ID="&ID)
		End If
		BBS.Execute("Delete from [sms] where MyFlag=2 And Flag=2")
		BBS.Execute("Update [User] set SmsSize=SmsSize-1 where ID="&BBS.MyID)
	End If
	Rs.close
	Set Rs=Nothing
	Response.Redirect "sms.asp"
End Sub

Sub DelAll
	Dim ID,I
	ID=BBS.CheckNum(request("ID"))
	I=0
	If ID=1 Then'ɾ����
		I=BBS.Execute("select count(*) From[Sms] where Name='"&MyName&"' And Flag=0" )(0)
		BBS.Execute("Update [sms] Set MyFlag=2 where MyName='"&BBS.MyName&"'")
	ElseIf ID=2 Then'ɾ����
		I=BBS.Execute("select count(*) From[Sms] where MyName='"&MyName&"' And Flag<>2" )(0)
		BBS.Execute("Update [sms] Set Flag=2 where Name='"&BBS.MyName&"'")
	Else
		BBS.Execute("Update [sms] Set Flag=2 where Name='"&BBS.MyName&"'")
		BBS.Execute("Update [sms] Set MyFlag=2 where MyName='"&BBS.MyName&"'")
	End If
	If isnull(I) Then I=0
	BBS.Execute("Update [User] set SmsSize="&i&" where ID="&BBS.MyID)
	BBS.Execute("Delete from [sms] where MyFlag=2 And Flag=2")
	Session(CacheName & "MyInfo") = Empty
	BBS.Alert"��������ż��ɹ���","sms.asp"
End Sub

Sub ShowMySmsInfo()
	Dim SmsSize,content
	SmsSize=int(SESSION(CacheName & "MyInfo")(20))
	AllSmsSize=SmsSize/Int(SESSION(CACHENAME & "MYGRADEINFO")(18))*100
	If AllSmsSize>100 Then AllSmsSize=100
	IF AllSmsSize<0 Then AllSmsSize=0
	IF AllSmsSize>0 And AllSmsSize<1 Then AllSmsSize=1
	Content=SmsSize/SESSION(CacheName & "MyGradeInfo")(18)*250
	If Content>250 Then Content=250
	Content="<div style='padding:3px;'><div style=""float:right;""><div style=""float:left; width:auto"">����������</div><div style=""float:left;width:250px;height:12px;border:#CCCCCC 1px dotted; background:#CCFFFF""><img src='Images/icon/hr1.gif' width='"&Content&"' height='12'></div>��ʹ�� <span style='color:#F00'>"&Int(AllSmsSize)&" </span>%</div><a href='sms.asp'><img src='Images/Icon/sms.gif' width='16' height='16' border='0' /> �ռ���</a> <a href='sms.asp?action=elapse'><img src='Images/icon/elapse.gif' border='0' /> ������</a> <a href='?action=write'><img border='0' src='Images/icon/add.gif' align=absmiddle> д������</a>&nbsp;<a href='#this' onclick=""if(confirm('��ȷ�����������������ż�����\n\n��ȷ��Ҫɾ����'))window.location.href='?action=delall'"" ><img src='Images/Icon/recycle.gif' border='0' align=absmiddle> �������</a></div>"
	Response.Write BBS.ReadSkins("�û��������")
	BBS.ShowTable"��̳��������",Content
End Sub



Sub ReadSms()
	Dim S,div,Content,Temp,UserPic,Rs,P,strPageInfo,Arr_Rs,I,Caption,bgColor,IUBB,Sqlwhere,title,UserName
	If action="elapse" Then
		Title="���͵��ż���¼"
		Sqlwhere="Name='"&BBS.MyName&"' and Flag=0"
	ElseIf action="colloquy" Then
		UserName=Request.querystring("Name")
		If Not BBS.Fun.CheckName(UserName) Then BBS.GoToErr(1)
		Title="��"&UserName&"�Ľ�̸��¼"
		Sqlwhere="(MyName='"&BBS.MyName&"' and Name='"&UserName&"' and MyFlag<2) or (Name='"&BBS.MyName&"' And  MyName='"&UserName&"' and Flag=0)"
	Else
		Title="��ȡ�ż�"
		Sqlwhere="MyName='"&BBS.MyName&"' and MyFlag<2"
	End If
	Set P = New Cls_PageView
	P.strTableName = "[Sms]"
	P.strPageUrl="?action="&action
	P.strFieldsList = "ID,Name,Content,AddTime,MyFlag,UbbString,Flag,MyName"
	P.strCondiction = Sqlwhere
	P.strOrderList = "ID desc"
	P.strPrimaryKey = "ID"
	P.intPageSize = 10
	P.intPageNow = Request.querystring("page")
	P.strCookiesName = "Sms"&action
	P.InitClass
	Arr_Rs = P.arrRecordInfo
	strPageInfo = P.strPageInfo
	Set P = nothing
	If IsArray(Arr_Rs) Then
		Set IUBB=New Cls_IUBB
		Div="<div style=""margin-left : 170px;min-height:150px;padding:10px;font-size:9pt;line-height:normal;word-wrap : break-word ;word-break : break-all ;border-left: 1px solid "&BBS.SkinsPIC(0)&""" onload=""this.style.overflowX='auto';"">"
		If BBS.MSIE Then
			Div=Replace(Div,"min-","padding-right:0px; overflow-x: hidden;")
		End If
		For i = 0 to UBound(Arr_Rs, 2)
			IUBB.UbbString=Arr_Rs(5,I)
			If lcase(Arr_Rs(1,I))=lcase(BBS.MyName) Then
				Temp="���͸� <a href='UserInfo.asp?Name="&Arr_Rs(7,I)&"'><strong>"&Arr_Rs(7,I)&"</strong></a> ���ż�&nbsp; "
				If action="elapse" Then Temp=Temp&"<a href='?action=colloquy&name="&Arr_Rs(7,I)&"'><img src='Images/icon/book.gif' border='0' alt='�鿴�Ự��¼' title='�鿴�Ự��¼' /></a> "
				If Session(CacheName & "MyInfo")(11)="1" Then
					UserPic="<img src='http://qqshow-user.tencent.com/"&Session(CacheName & "MyInfo")(10)&"/11/' alt='QQͷ��' />"
				Else
					UserPic="<img src="&Session(CacheName & "MyInfo")(12)&" width="&Session(CacheName & "MyInfo")(13)&" height="&Session(CacheName & "MyInfo")(14)&" alt='' />"
				End if
			Else
				Set Rs=BBS.execute("select top 1 IsQQpic,QQ,Pic,PicW,PicH from [User] where Name='"&Arr_Rs(1,I)&"'")
				 If Not Rs.eof then
					Temp="<a href='UserInfo.asp?Name="&Arr_Rs(1,I)&"'><img border='0' src='Images/icon/info.GIF' alt='�鿴����' /></a> <a href='?action=write&Name="&Arr_Rs(1,I)&"&id="&Arr_Rs(0,I)&"'><img border='0' src='Images/icon/reply.gif' alt='�ظ�' title='�ظ�' /></a> <a href='?action=colloquy&name="&Arr_Rs(1,I)&"'><img src='Images/icon/book.gif' border='0' alt='�鿴�Ự��¼' title='�鿴�Ự��¼' /></a> "
					IF Rs(0)=1 then
						UserPic="<img src='http://qqshow-user.tencent.com/"&Rs(1)&"/11/' alt='QQͷ��' />"
					Else
						UserPic="<img border='0' src='"&rs(2)&"' width='"&rs(3)&"' height='"&rs(4)&"' alt='' />"
					End If
				End if
				Rs.Close
				Set Rs=nothing
			End If
			
			If I mod 2 <>0 Then bgColor="background-color: "&BBS.SkinsPIC(1)&";" Else bgColor=""
			S="<div style="""&bgColor&";text-align:left;border-bottom:1px solid "&BBS.SkinsPIC(0)&"""><div style='float:left;text-indent:24px;width:170px'><br /><div><b>"&Arr_Rs(1,I)&"</b></div><div>"&UserPic&"</div></div>"
			S=S&DIV&Temp&"<a href='#this' onclick=""if(confirm('��ȷ����ɾ���������ԣ���\n\n��ȷ��Ҫɾ����'))window.location.href='?id="&Arr_Rs(0,I)&"&action=del'"" ><img border='0' alt='ɾ��' src='Images/Icon/delete.gif' /></a> "
			IF Arr_Rs(4,I)=1 Then S=S&"<img src='Images/Icon/New.Gif' alt='�µ�����' />"
			S=S&"<hr width='98%' size='1' color="""&BBS.SkinsPIC(0)&""" ><blockquote>"&IUBB.UBB(Arr_Rs(2,I),2)&"<p></p><div align=""right""><img src='Images/icon/add.gif' border='0' atl='' /> ����ʱ�䣺 "&Arr_Rs(3,I)&"</div></blockquote></div></div>"
			Content=Content&S
		Next
		Content=Content&"<div style=""padding:5px;BACKGROUND: "&BBS.SkinsPIC(2)&";"">"&strPageInfo&"<br><br></div>"
		Set IUBB=Nothing
	End If
	BBS.ShowTable Title,Content
	If Session(CacheName&"updateSms")="" or Int(Session(CacheName & "MyInfo")(27))>0 then
		BBS.ExeCute("Update [user] Set NewSmsNum=0 Where Name='"&BBS.MyName&"'")
		BBS.ExeCute("Update [Sms] Set MyFlag=0 Where MyFlag=1 and MyName='"&BBS.MyName&"'")
		Session(CacheName&"updateSms")="Y"
		Session(CacheName & "MyInfo") = Empty
	End If
End Sub

Sub WriteSms()
	If AllSmsSize=100 Then
		Temp="ϵͳ����":S="<br><P>&nbsp;&nbsp;�װ����û���������̳�������������������뾡��ɾ��һЩ�ż���</p><br>"
	Else
		Dim Name,Rs,S,Temp,Content,ID
		ID=BBS.CheckNum(request("ID"))
		Name=request.querystring("Name")
		If Not BBS.Fun.CheckName(Name) Then BBS.GoToErr(1)
		Set Rs=BBS.execute("select Content from [sms] where name='"&Name&"' And MyName='"&BBS.MyName&"' and Id="&ID&"")
		if not Rs.eof then 
		Content=Rs("Content")
		End if
		Rs.Close
		Set Rs=nothing
		S="<form style='margin:0;' method='POST' action='?action=save' name='say'>"
		S=S&BBS.Row("<b>���Զ���</b>","<textarea id='content' name='content' style='display:none'>"&Content&"</textarea><input type=hidden name='iCode' id='iCode' value='BBS' /><input name='caption' type='text' class='text' id='caption' size='30' value='"&Name&"'>","75%","")
		Temp="<b>�ż����ݣ�</b><br /> <a href=""javascript:CheckLength("&Session(CacheName & "MyGradeInfo")(19)&")"">��������"&Session(CacheName & "MyGradeInfo")(19)&"���ֽ�</a><br />"
		Temp=Temp&"ÿ�������Է���"&Session(CacheName & "MyGradeInfo")(13)&"��"
		If Int(BBS.Info(123)) >0 Then Temp=Temp & "<br />ÿ����ȡ���ͷѣ�"&BBS.Info(123)&BBS.Info(120)
	If BBS.Info(60)="1" Then Content="UbbEdit()" Else Content="HtmlEdit()"
	Content="<script type=""text/javascript"">"&Content&"</script>"
	S=S&BBS.Row(""&Temp,Content,"75%","")
	S=S&"<div align='center' style=""padding:5px;BACKGROUND: "&BBS.SkinsPIC(1)&";"">"
	S=S&"&nbsp;<input type='button' value=' �� �ͣ�' id='sayb' onclick='checkform("&Session(CacheName & "MyGradeInfo")(19)&")' class='button' /> <input type='reset' value=' �� д ' onclick='Goreset()' class='button' />" 
	S=S&"</div></form>"
	Temp="ǩд��������"
	End If
	BBS.ShowTable Temp,S
End Sub

Sub SaveSms()
	'BBS.CheckMake()
	Dim S,Content,ToName,TmpUbbString
	If int(SESSION(CacheName & "MyInfo")(7))<int(BBS.Info(123)) Then BBS.GoToErr(52)
	If Session(CacheName&"SmsTime")+1/1440>now() then BBS.GoToErr(53)
	ToName=BBS.Fun.GetForm("caption")
	Content=BBS.Fun.GetForm("Content")
	If ToName="" or Content=""  then BBS.GoToErr(36)
	If BBS.Fun.CheckIsEmpty(Content) Then BBS.GoToErr(50)
	If BBS.Info(60)="1" Then Content=BBS.Fun.Replacehtml(Content)
	TmpUbbString=BBS.Fun.UbbString(Content)
	If Not BBS.Fun.CheckName(ToName) Then BBS.GoToErr(41)
	IF Len(Content)>Int(Session(CacheName & "MyGradeInfo")(19)) Then BBS.GoToErr(29)
	S=BBS.Execute("Select Count(*) From[Sms] where Name='"&BBS.MyName&"' And DATEDIFF('d',AddTime,'"&BBS.NowBbsTime&"')<1")(0)
	If Isnull(S) Then S=0
	If S>Int(Session(CacheName & "MyGradeInfo")(13)) Then BBS.GoToErr(55)
	If BBS.execute("select Name From [User] where name='"&ToName&"'and IsDel=0").eof Then BBS.GoToErr(54)
	BBS.execute("insert into [sms](name,Content,Myname,ubbString,MyFlag)values('"&BBS.MyName&"','"&Content&"','"&ToName&"','"&TmpUbbString&"',1)")
	BBS.execute("update [user] Set Coin=Coin-"&int(BBS.Info(123))&" where ID="&BBS.MyID)
	BBS.ExeCute("Update [user] Set NewSmsNum=NewSmsNum+1,SmsSize=SmsSize+1 Where Name='"&ToName&"'")
	Session(CacheName&"SmsTime")=Now()
	'����֪ͨ
	BBS.UpdageOnline ToName,1
	Content="<div style='margin:15px;line-height:150%'><li>�Ѿ��ɹ��ĸ� <b>"&ToName&"</b> ����</li><li>��վ�۳������� "&BBS.Info(123)&BBS.Info(120)&"</li><li><a href=""index.asp"">������ҳ</a> </li><li><a href=""sms.asp"">�����ҵ�����</a></li></Div>"
	BBS.ShowTable"���ͳɹ�",Content
End Sub
%>