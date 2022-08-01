<!-- #include file="Inc.asp" -->
<!-- #include file="Inc/Page_Cls.asp" -->
<%Dim ID,Rs,Page_Url
If Not BBS.Founduser Then BBS.GoToErr(10)
If Request.QueryString("page") > 1 Then
  Page_Url = "?Page="&Request.QueryString("page")
Else
  Page_Url = ""
End If
BBS.Head"faction.asp"&Page_Url,"","论坛帮派"
ID = BBS.CheckNum(Request.querystring("ID"))
Select Case Left(Request("Action"),10)
Case"Add"
	Add
Case"Edit"
	Edit
Case"FactionAdd"
	FactionAdd
Case"FactionOut"
	FactionOut
Case"Look"
	Look
Case"Del"
	Del
Case Else
	Main()
End Select
BBS.Footer
Set BBS =Nothing


Sub Main()
Dim intPageNow,strPageInfo,arr_Rs,i,Pages,page,Content
Content="<table width='95%' border=0 style='border-collapse: collapse;WORD-BREAK: break-all;'><tr><td align='center' height=40 colspan=5><a href='?action=Add'><img src='Images/icon/right.gif' border='0' align='absmiddle'> 创建帮派</a>&nbsp;&nbsp;<a href=#this onclick=""if(confirm('您确定要退出该帮派？\n\n您的"&BBS.Info(121)&"将会减去 1'))window.location.href='?action=FactionOut'""><img src='Images/icon/right.gif'  border='0' align='absmiddle'> 退出帮派</a></td></tr>"&_
"<tr><td width='15%' class=FactionTit>派别</td><td width='40%' class=FactionTit>宗旨</td><td width='15%' class=FactionTit>创始人</td><td width='10%' class=FactionTit>动作</td><td width='20%' class=FactionTit>帮主管理</td></tr>"
	intPageNow = Request.QueryString("page")
	Set Pages = New Cls_PageView
	Pages.strTableName = "[Faction]"
	Pages.strFieldsList = "ID,Name,Note,User,BuildDate"
	Pages.strOrderList = "ID desc"
	Pages.strPrimaryKey = "ID"
	Pages.intPageSize = 15
	Pages.intPageNow = intPageNow
	Pages.strCookiesName = "Faction"'客户端记录总数
	Pages.Reloadtime=3'每三分钟更新Cookies
	Pages.strPageVar = "page"
	Pages.InitClass
	Arr_Rs = Pages.arrRecordInfo
	strPageInfo = Pages.strPageInfo
	Set Pages = nothing
	If IsArray(Arr_Rs) Then
	For i = 0 to UBound(Arr_Rs, 2)
		Content=Content & "<tr><td align='center' height='25'><a href=?action=Look&ID="&Arr_Rs(0,i)&">"&Arr_Rs(1,i)&"</a></td><td>"&Arr_Rs(2,i)&"</td><td align='center' height='25'><a href=UserInfo.asp?Name="&Arr_Rs(3,i)&">"&Arr_Rs(3,i)&"</a></td><td align='center'>"
		If SESSION(CacheName & "MyInfo")(25)=Arr_Rs(1,i) then
			Content=Content & "<a href=#this onclick=""if(confirm('您确定要退出该帮派？\n\n您的"&BBS.Info(121)&"将会减去 1'))window.location.href='?action=FactionOut&ID="&Arr_Rs(0,i)&"'"">退出此帮</a>"
		Else
			Content=Content & "<a href=#this onclick=""if(confirm('您确定要加入该帮派？\n\n您的"&BBS.Info(121)&"必须达到 3'))window.location.href='?action=FactionAdd&ID="&Arr_Rs(0,i)&"'"">加入此帮</a>"
		End if
		Content=Content & "<td align='center'><a href='?action=Edit&ID="&Arr_Rs(0,i)&"'><img src='Images/icon/edit.gif' border='0' alt='' />修改</a> <a href=#this onclick=""if(confirm('您确定要解散该帮派？'))window.location.href='?action=Del&ID="&Arr_Rs(0,i)&"'""><img src='Images/icon/del.gif' border='0' alt='' />解散</a></td></tr>"
	Next
	End If
	Content=Content & "</table><div style=""padding:5px;BACKGROUND: "&BBS.SkinsPIC(2)&";"">"&strPageInfo&"<br><br></div>"

	BBS.ShowTable"论坛帮派",Content
End Sub

Sub FactionAdd
	Dim Content,Rs
	BBS.CheckMake
	If SESSION(CacheName & "MyInfo")(25)<>"" Then
		BBS.Alert"您已经加入["&SESSION(CacheName & "MyInfo")(25)&"]了！请先退出["&SESSION(CacheName & "MyInfo")(25)&"]才能加入新帮","?"
	ElseIf Int(SESSION(CacheName & "MyInfo")(6))<3 then
		BBS.Alert"您的"&BBS.Info(121)&"值小于3！没有资格加入帮派！","?"
	Else
		Set Rs=BBS.Execute("select Name from [Faction] where ID="&ID)
		IF Not Rs.Eof Then
			BBS.Execute("update [user] Set Faction='"&rs(0)&"' where Name='"&BBS.MyName&"'")
			Session(CacheName & "MyInfo") = Empty
			BBS.Alert"成功的加入 ["&Rs(0)&"] 帮派！","?"
		Else
			BBS.Alert"不存在这个帮派！","?"
		End If
		Rs.Close
	End If
End Sub

Sub FactionOut
	BBS.CheckMake
	If SESSION(CacheName & "MyInfo")(25)="" Then
		BBS.Alert"您目前还没有加入任何帮派！","?"
	Else
		If Not BBS.Execute("select ID from [Faction] where user='"&BBS.MyName&"'").eof Then
			BBS.Alert"您是掌门人，不能退出帮派！退出必需先要解散帮派！","?"
		Else
			BBS.execute("Update [user] Set Faction='',Mark=Mark-1 where name='"&BBS.MyName&"'")
			Session(CacheName & "MyInfo") = Empty
		End If
		BBS.Alert"退出帮派成功","?"
	End If
End Sub

Sub Del
BBS.CheckMake
	Set Rs=BBS.Execute("Select Name,User From[Faction] where ID="&ID)
	If Rs.Eof Then
		BBS.Alert"不存在这个帮派！","?"
	ElseIf BBS.MyName<>Rs(1) Then
		BBS.Alert"您不是该帮的帮主无法解散该帮！","?"
	Else
		BBS.Execute("Update [user] set Faction='' where Faction='"&rs(0)&"'")
		BBS.Execute("Delete from [Faction] where ID="&ID)
		Session(CacheName & "MyInfo") = Empty
		BBS.Alert"解散帮派成功！","?"
	End if
	Rs.Close
End Sub

Sub Look
Dim Content
Set Rs=BBS.Execute("Select Name,FullName,Note,User,BuildDate from [Faction] where ID="&ID)
If Rs.eof Then
	BBS.Alert"不存在此帮派！","?"
Else
	Content="<table width='95%' border=0 style='border-collapse: collapse;WORD-BREAK: break-all;'><tr><td width='24%' align='right' height=25><b>帮派名称：</b></td><td width='74%'>&nbsp;"&BBS.Fun.HtmlCode(rs(0))&"</td></tr>"&_
	"<tr><td align='right' height=25><b>帮派全称：</b></td><td>&nbsp;"&BBS.Fun.HtmlCode(rs(1))&"</td></tr>"&_
	"<tr><td align='right' height=25><b>帮派宗旨：</b></td><td>&nbsp;"&BBS.Fun.HtmlCode(rs(2))&"</td></tr>"&_
	"<tr><td align='right' height=25><b>创建时间：</b></td><td>&nbsp;"&Rs(4)&"</td></tr>"&_
	"<tr><td align='right' height=25><b>帮主名称：</b></td><td>&nbsp;"&Rs(3)&"</td></tr>"&_
	"<tr><td align='right' height=25><b>现有弟子：</b></td><td>"&Desciple(Rs(0))&"</td></tr>"&_
	"<tr><td colspan=2 align='center' height=25><a href='?'>【返回】</a></td></tr></table>"
	BBS.ShowTable"帮派信息",Content
End If
Rs.Close
End Sub

Function Desciple(name)
	Dim dRs,I
	I=0
	Set dRs=BBS.Execute("Select Name From [user] where Faction='"&Name&"'")
	Do while not dRs.eof
	I=I+1
	Desciple=Desciple & "<a target='_blank' href='Userinfo.asp?Name="&dRs(0)&"'>"&dRs(0)&"</a>&nbsp;&nbsp;&nbsp;"
	dRs.movenext
	Loop
	dRs.close:Set dRs=NoThing
	Desciple="<table width='95%' border='0' cellpadding='0' cellspacing='0'><tr><td>&nbsp;"&I&" 名</td><td width='90%'><marquee onmouseover='this.stop()' onmouseout='this.start()' scrollAmount='3' direction='left' width='95%' height='15'>"&Desciple&"</marquee></td></tr></table>"
End Function

Sub Add
Dim Name,FullName,Note,Content
BBS.CheckMake
Name=BBS.Fun.GetStr("Name")
FullName=BBS.Fun.GetStr("FullName")
Note=BBS.Fun.GetStr("Note")
IF Name="" And FullName="" And Note="" Then
	Content="<form  method='post' style='margin:0'>"&_
	"<table width='95%' border=0 style='border-collapse: collapse;WORD-BREAK: break-all;'><tr><td colspan=2 height=30 align='center'><font color=red>创建门派的必要条件： 1.您的 "&BBS.Info(121)&" 必须 20 以上！ 2.需要扣除您 10000 的"&BBS.Info(120)&"作为门派基金！ </font></td></tr>"&_
	"<tr><td align='right' height=25><b>帮派名称：</b></td><td>&nbsp;<input class='text' Maxlength=10 Name='Name' size='10'>*最多只能6个汉字</td></tr>"&_
	"<tr><td align='right' height=25><b>帮派全称：</b></td><td>&nbsp;<input class='text' size=30 name='FullName'> * </td></tr>"&_
	"<tr><td align='right' height=25><b>帮派宗旨：</b></td><td>&nbsp;<input class='text' size=70 name='Note'> * </td></tr>"&_
	"</table><div align='center' style=""height:25px;BACKGROUND: "&BBS.SkinsPIC(2)&";""><input type='submit' class='button' value=' 创 建 '>&nbsp;&nbsp;<input type='reset' class='button' value=' 重 填 '></div></form>"
	BBS.ShowTable"创建帮派",Content
Else
	IF Name="" or FullName="" or Note="" Then
		BBS.Alert"帮派要填写的信息你没有填写完整。","?"
	ElseIF Len(Name)>6 or Len(FullName)>50 Or Len(Note)>200 Then
		BBS.Alert"字符太多，超过了论坛的限制。","?"
	ElseIf int(SESSION(CacheName & "MyInfo")(6))<20 then
		BBS.Alert"您的"&BBS.Info(121)&"小于 20 ！","?"
	ElseIf int(SESSION(CacheName & "MyInfo")(7))<10000 then
		BBS.Alert"您的"&BBS.Info(120)&"少于 10000 ！","?"
	ElseIf Not BBS.Execute("Select ID From[Faction] where User='"&BBS.MyName&"'").Eof Then
		BBS.Alert"您已经贵为帮主了，不能再创立帮派！","?"
	Else
	BBS.execute("Insert into[Faction](Name,FullName,[Note],BuildDate,[User])Values('"&Name&"','"&FullName&"','"&Note&"','"&BBS.NowBbsTime&"','"&BBS.MyName&"')")
	BBS.execute("Update [User] Set Coin=Coin-10000,Faction='"&Name&"' where ID="&BBS.MyID&"")
	Session(CacheName & "MyInfo") = Empty
	BBS.Alert"您成功的创建了帮派["&Name&"]，您现在是该帮派的掌门人！恭喜您！","?"
	End if
End if
End Sub

Sub Edit
Dim Name,FullName,Note,Content
Name=BBS.Fun.GetStr("Name")
FullName=BBS.Fun.GetStr("FullName")
Note=BBS.Fun.GetStr("Note")
Set Rs=BBS.Execute("Select Name,FullName,Note,User from [Faction] where ID="&ID)
If Rs.eof Then
	BBS.Alert"不存在此帮派！","?"
ElseIf BBS.MyName<>Rs(3) Then
	BBS.Alert"您不是["&Rs(0)&"]的帮主无法修改信息！","?"
Else
	IF Name="" And FullName="" And Note="" Then
		Set Rs=BBS.Execute("Select Name,FullName,Note,User from [Faction] where ID="&ID)
		If Rs.eof Then
			BBS.Alert"不存在此帮派！","?"
		ElseIf BBS.MyName<>Rs(3) Then
			BBS.Alert"您不是["&Rs(0)&"]的帮主无法修改信息！","?"
		Else
			Content="<form  method='post' style='margin:0'>"&_
			"<table width='95%' border=1' style='border-collapse: collapse;WORD-BREAK: break-all;'><tr><td colspan=2 height=30 align='center'><font color=red>注意：每次修改帮派信息将扣除您 1000 的"&BBS.Info(120)&"！ </font></td></tr>"&_
			"<tr><td align='right' height=25><b>帮派名称：</b></td><td>&nbsp;<input class='text' Maxlength=10 Name='Name' size='10' value='"&Rs(0)&"'>*不要超过5个汉字</td></tr>"&_
			"<tr><td align='right' height=25><b>帮派全称：</b></td><td>&nbsp;<input class='text' size=30 name='FullName' value='"&Rs(1)&"'> * </td></tr>"&_
			"<tr><td align='right' height=25><b>帮派宗旨：</b></td><td>&nbsp;<input class='text' size=70 name='Note' value='"&Rs(2)&"'> * </td></tr>"&_
			"</table><div align='center' style=""height:25px;BACKGROUND: "&BBS.SkinsPIC(2)&";""><input type='submit' class='button' value=' 修 改 '>&nbsp;&nbsp;<input type='reset' class='button' value=' 重 填 '></div></form>"
			BBS.ShowTable"创建帮派",Content
		End If
	Else
		IF Name="" or FullName="" or Note="" Then
			BBS.Alert"帮派要填写的信息你没有填写完整。","?"
		ElseIf int(SESSION(CacheName & "MyInfo")(7))<1000 then
			BBS.Alert"对不起，你的"&BBS.Info(120)&"少于1000元，不能整顿帮派。","?"
		ElseIF Len(Name)>10 or Len(FullName)>50 Or Len(Note)>200 Then
			BBS.Alert"字符太多，超过了论坛的限制。","?"
		Else
		BBS.execute("Update [User] Set Faction='"&Name&"' where Faction='"&Rs(0)&"'")
		BBS.execute("Update [User] Set Coin=Coin-1000 where Name='"&BBS.MyName&"'")
		BBS.execute("Update [Faction]Set Name='"&Name&"',FullName='"&FullName&"',[Note]='"&Note&"' where ID="&ID)
		Session(CacheName & "MyInfo") = Empty
		BBS.Alert"成功的修改了帮派！","?"
		End if
	End if
End If
Rs.Close
End Sub
%>