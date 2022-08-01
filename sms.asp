<!--#include file="Inc.asp"-->
<!--#include file="Inc/Page_Cls.asp"-->
<!--#include file="inc/ubb_Cls.asp"-->
<%
Dim action,AllSmsSize
If Not BBS.FoundUser Then BBS.GoToErr(20)
BBS.Position=BBS.Position&" -> <a href=""userinfo.asp"">用户控制面版</a>"
action=Lcase(Request("action"))
BBS.Head "sms.asp","","处理信件"
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
	If ID=1 Then'删收箱
		I=BBS.Execute("select count(*) From[Sms] where Name='"&MyName&"' And Flag=0" )(0)
		BBS.Execute("Update [sms] Set MyFlag=2 where MyName='"&BBS.MyName&"'")
	ElseIf ID=2 Then'删发箱
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
	BBS.Alert"清空所有信件成功！","sms.asp"
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
	Content="<div style='padding:3px;'><div style=""float:right;""><div style=""float:left; width:auto"">信箱容量：</div><div style=""float:left;width:250px;height:12px;border:#CCCCCC 1px dotted; background:#CCFFFF""><img src='Images/icon/hr1.gif' width='"&Content&"' height='12'></div>已使用 <span style='color:#F00'>"&Int(AllSmsSize)&" </span>%</div><a href='sms.asp'><img src='Images/Icon/sms.gif' width='16' height='16' border='0' /> 收件箱</a> <a href='sms.asp?action=elapse'><img src='Images/icon/elapse.gif' border='0' /> 发件箱</a> <a href='?action=write'><img border='0' src='Images/icon/add.gif' align=absmiddle> 写新留言</a>&nbsp;<a href='#this' onclick=""if(confirm('按确定将清空邮箱的所有信件！！\n\n您确定要删除吗？'))window.location.href='?action=delall'"" ><img src='Images/Icon/recycle.gif' border='0' align=absmiddle> 清空信箱</a></div>"
	Response.Write BBS.ReadSkins("用户控制面版")
	BBS.ShowTable"论坛留言信箱",Content
End Sub



Sub ReadSms()
	Dim S,div,Content,Temp,UserPic,Rs,P,strPageInfo,Arr_Rs,I,Caption,bgColor,IUBB,Sqlwhere,title,UserName
	If action="elapse" Then
		Title="发送的信件记录"
		Sqlwhere="Name='"&BBS.MyName&"' and Flag=0"
	ElseIf action="colloquy" Then
		UserName=Request.querystring("Name")
		If Not BBS.Fun.CheckName(UserName) Then BBS.GoToErr(1)
		Title="和"&UserName&"的交谈记录"
		Sqlwhere="(MyName='"&BBS.MyName&"' and Name='"&UserName&"' and MyFlag<2) or (Name='"&BBS.MyName&"' And  MyName='"&UserName&"' and Flag=0)"
	Else
		Title="收取信件"
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
				Temp="发送给 <a href='UserInfo.asp?Name="&Arr_Rs(7,I)&"'><strong>"&Arr_Rs(7,I)&"</strong></a> 的信件&nbsp; "
				If action="elapse" Then Temp=Temp&"<a href='?action=colloquy&name="&Arr_Rs(7,I)&"'><img src='Images/icon/book.gif' border='0' alt='查看会话记录' title='查看会话记录' /></a> "
				If Session(CacheName & "MyInfo")(11)="1" Then
					UserPic="<img src='http://qqshow-user.tencent.com/"&Session(CacheName & "MyInfo")(10)&"/11/' alt='QQ头像' />"
				Else
					UserPic="<img src="&Session(CacheName & "MyInfo")(12)&" width="&Session(CacheName & "MyInfo")(13)&" height="&Session(CacheName & "MyInfo")(14)&" alt='' />"
				End if
			Else
				Set Rs=BBS.execute("select top 1 IsQQpic,QQ,Pic,PicW,PicH from [User] where Name='"&Arr_Rs(1,I)&"'")
				 If Not Rs.eof then
					Temp="<a href='UserInfo.asp?Name="&Arr_Rs(1,I)&"'><img border='0' src='Images/icon/info.GIF' alt='查看资料' /></a> <a href='?action=write&Name="&Arr_Rs(1,I)&"&id="&Arr_Rs(0,I)&"'><img border='0' src='Images/icon/reply.gif' alt='回复' title='回复' /></a> <a href='?action=colloquy&name="&Arr_Rs(1,I)&"'><img src='Images/icon/book.gif' border='0' alt='查看会话记录' title='查看会话记录' /></a> "
					IF Rs(0)=1 then
						UserPic="<img src='http://qqshow-user.tencent.com/"&Rs(1)&"/11/' alt='QQ头像' />"
					Else
						UserPic="<img border='0' src='"&rs(2)&"' width='"&rs(3)&"' height='"&rs(4)&"' alt='' />"
					End If
				End if
				Rs.Close
				Set Rs=nothing
			End If
			
			If I mod 2 <>0 Then bgColor="background-color: "&BBS.SkinsPIC(1)&";" Else bgColor=""
			S="<div style="""&bgColor&";text-align:left;border-bottom:1px solid "&BBS.SkinsPIC(0)&"""><div style='float:left;text-indent:24px;width:170px'><br /><div><b>"&Arr_Rs(1,I)&"</b></div><div>"&UserPic&"</div></div>"
			S=S&DIV&Temp&"<a href='#this' onclick=""if(confirm('按确定将删除这条留言！！\n\n您确定要删除吗？'))window.location.href='?id="&Arr_Rs(0,I)&"&action=del'"" ><img border='0' alt='删除' src='Images/Icon/delete.gif' /></a> "
			IF Arr_Rs(4,I)=1 Then S=S&"<img src='Images/Icon/New.Gif' alt='新的留言' />"
			S=S&"<hr width='98%' size='1' color="""&BBS.SkinsPIC(0)&""" ><blockquote>"&IUBB.UBB(Arr_Rs(2,I),2)&"<p></p><div align=""right""><img src='Images/icon/add.gif' border='0' atl='' /> 留言时间： "&Arr_Rs(3,I)&"</div></blockquote></div></div>"
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
		Temp="系统警告":S="<br><P>&nbsp;&nbsp;亲爱的用户，您的论坛留言信箱容量已满，请尽快删除一些信件！</p><br>"
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
		S=S&BBS.Row("<b>留言对象：</b>","<textarea id='content' name='content' style='display:none'>"&Content&"</textarea><input type=hidden name='iCode' id='iCode' value='BBS' /><input name='caption' type='text' class='text' id='caption' size='30' value='"&Name&"'>","75%","")
		Temp="<b>信件内容：</b><br /> <a href=""javascript:CheckLength("&Session(CacheName & "MyGradeInfo")(19)&")"">内容限制"&Session(CacheName & "MyGradeInfo")(19)&"个字节</a><br />"
		Temp=Temp&"每天最多可以发送"&Session(CacheName & "MyGradeInfo")(13)&"封"
		If Int(BBS.Info(123)) >0 Then Temp=Temp & "<br />每封收取发送费："&BBS.Info(123)&BBS.Info(120)
	If BBS.Info(60)="1" Then Content="UbbEdit()" Else Content="HtmlEdit()"
	Content="<script type=""text/javascript"">"&Content&"</script>"
	S=S&BBS.Row(""&Temp,Content,"75%","")
	S=S&"<div align='center' style=""padding:5px;BACKGROUND: "&BBS.SkinsPIC(1)&";"">"
	S=S&"&nbsp;<input type='button' value=' 发 送！' id='sayb' onclick='checkform("&Session(CacheName & "MyGradeInfo")(19)&")' class='button' /> <input type='reset' value=' 重 写 ' onclick='Goreset()' class='button' />" 
	S=S&"</div></form>"
	Temp="签写发送留言"
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
	'在线通知
	BBS.UpdageOnline ToName,1
	Content="<div style='margin:15px;line-height:150%'><li>已经成功的给 <b>"&ToName&"</b> 留言</li><li>本站扣除手续费 "&BBS.Info(123)&BBS.Info(120)&"</li><li><a href=""index.asp"">返回首页</a> </li><li><a href=""sms.asp"">返回我的信箱</a></li></Div>"
	BBS.ShowTable"发送成功",Content
End Sub
%>