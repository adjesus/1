<!--#include file="admin_Check.asp"-->
<%
Head()
Select Case Lcase(request.querystring("action"))
Case"bbsinfo"
	CheckString "01"
	BbsInfo
Case"updateconfigdata"
	CheckString "02"
	UpdateConfigData
Case"saveplacard"
	CheckString "03"
	Placard
Case"delplacard"
	CheckString "03"
	delPlacard
Case"gapad"
	CheckString "04"
	Gapad
Case"dellink"
	CheckString "05"
	DelLink
Case"savelink"
	CheckString "05"
	SaveLink
Case"updatelink"
	CheckString "05"
	UpdateLink
Case"islockip"
	CheckString "06"
	IsLockIP
Case"lockip"
	CheckString "06"
	LockIP	
Case"dellockip"
	CheckString "06"
	DelLockIP
Case"dellog"
	CheckString "07"
	DelLog
Case "clean"
	CheckString "08"
	Clean
Case"topadmin"
	CheckString "22"
	Topadmin
Case"boardadmin"
	CheckString "23"
	Boardadmin
Case"allboardadmin"
	CheckString "23"
	AllBoardadmin
Case"allupdategrade"
	CheckString "27"
	AllUpdateGrade
Case"delgrade"
	CheckString "27"
	DelGrade
Case"savegrade"
	CheckString "27"
	SaveGrade
Case"delessay"
	CheckString "31"
	DelEssay
Case"moveessay"
	CheckString "32"
	MoveEssay
Case"delsms"
	CheckString "33"
	DelSms
Case"allsms"
	CheckString "34"
	AllSms
Case"savemenu"
	CheckString "41"
	SaveMenu
Case"delmenu"
	CheckString "41"
	DelMenu
Case"menuorder"
	CheckString "41"
	MenuOrder
Case"setjsmenu"
	'CheckString ""
	SetJsMenu
Case"saveagreement"
	CheckString "42"
	saveagreement
Case "bank"
	CheckString "44"
	Bank	
Case "delfaction"
	CheckString "45"
	DelFaction
Case "savefaction"
	CheckString "45"
	SaveFaction
Case"compressdata"
	CheckString "51"
	CompressData
Case"notcompactdb"
	NotCompactDB
Case"okcompactdb"
	OkCompactdb
Case"backupdata"
	CheckString "52"
	BackupData
Case"restoredata"
	CheckString "53"
	RestoreData
Case"autesqltable"
	CheckString "54"
	AuteSqlTable
Case"addsqltable"
	CheckString "54"
	AddSqlTable
Case"delsqltable"
	CheckString "54"
	DelSqlTable
Case"sqltableunite"
	CheckString "55"
	SqlTableUnite
Case"updatebbsdate"
	CheckString "55"
	UpdateBbsDate
Case"updatetopic"
	CheckString "55"
	UpdateTopic
Case"updatealluser"
	CheckString "55"
	UpdateAllUser
Case"delwuiong"
	CheckString "55"
	DelWuiong
End select
Footer()


Sub BbsInfo()
	Dim I,S
	For i=0 to 123
	If instr(",0,1,2,4,5,6,7,17,18,19,28,29,34,35,36,37,46,50,52,58,59,79,82,83,84,85,86,87,88,89,119,120,121,122,",","&i&",")<>0 Then 
	If i=119 then
	If not isnumeric(BBS.Fun.GetStr("info"&i)) Then GoBack"","银行利率请用数字填写！":Exit Sub
	End If
	Else
		IF BBS.Fun.GetStr("info"&i)="" Then GoBack"",i:Exit Sub
		If Not BBS.Fun.isInteger(BBS.Fun.GetStr("info"&i)) then
			GoBack"","一些参数必须填为正整数，否则论坛不能正常运行。"&i
			Exit Sub
		End if
	End IF
		S=S&Replace(Request.form("info"&i),",","&#44")&","
	Next
	S=S&",0,0,0,0,0"
	BBS.execute("update [Config] set Info='"&S&"'")
	S="成功更改论坛信息设置"
	BBS.NetLog"操作后台_"&S
	Suc"修改成功",S,"admin_action.asp?action=BbsInfo"
	BBS.Cache.Clean("Info")
End Sub

Sub UpdateConfigData()
	With BBS
	Dim S,UserNum,AllEssayNum,TopicNum,MaxEssayNum,MaxOnlineNum,MaxOnlineTime,Hits
	Hits=Request.form("hits")
	UserNum=Request.form("usernum")
	AllEssayNum=Request.form("allessaynum")
	TopicNum=Request.form("topicnum")
	MaxEssayNum=Request.form("maxessaynum")
	MaxOnlineNum=Request.form("maxonlinenum")
	MaxOnlineTime=.Fun.GetForm("maxonlinetime")
	If Not .Fun.isInteger(Hits) or Not .Fun.isInteger(UserNum) Or Not .Fun.isInteger(AllEssayNum) or Not .Fun.isInteger(TopicNum) or Not .Fun.isInteger(MaxEssayNum) or Not .Fun.isInteger(MaxOnlineNum) Then
		GoBack"","一些参数必须填为正整数，否则论坛不能正常运行。"
	End if
	.Execute("update [Config] set Hits="&hits&",UserNum="&UserNum&",AllEssayNum="&AllEssayNum&",TopicNum="&TopicNum&",MaxEssayNum="&MaxEssayNum&",MaxOnlineNum="&MaxOnlineNum&",MaxOnlineTime='"&MaxOnlineTime&"'")
	S="论坛系统数据统计修改成功"
	Suc"修改成功",S,"admin_action.asp?action=ConfigData"
	.NetLog"操作后台_"&S
	.Cache.Clean("InfoUpdate")
	.Cache.Clean("Hits")
	End With
End Sub

Sub DelLog()
	If Request.Form("Del")="清空日志" Then
		BBS.Execute("Delete From [Log] where DATEDIFF('d', LogTime,'"&BBS.NowBBSTime&"')>2")
	Else
	Dim ID
		ID=Request.form("ID")
		If ID="" Then Goback "","请先选择":Exit Sub
		BBS.Execute("Delete From [Log] where ID in("&ID&") And DATEDIFF('d', LogTime,'"&BBS.NowBBSTime&"')>2")
	End If
	BBS.NetLog"操作后台_日志系统-"&Request.Form("Del")
	Suc "","删除日志成功!系统会自动保留二天的日志记录。","admin_actionList.asp?action=Log"
End Sub

Sub UpdateUserList()
	Dim ID,point,AllTable,i,S,UserName,IsBe
	IsBe=False
	UserName=Request("Name")
	point=Request("point")
	ID=Request("ID")
	If UserName<>"" Then
		Set Rs=BBS.Execute("Select ID From [User] where Name='"&UserName&"'")
		If Rs.eof Then
		Goback"","该用户还没有注册！"
		Exit Sub
		Else
		ID=Rs(0)
		End If
		Rs.Close
	End If
	If ID="" Then Goback "","请先选择用户":Exit Sub
	If Point="" Then Goback"","你还没有选定如何进行操作":Exit Sub
	Set Rs=BBS.Execute("Select Name,IsVIP,IsDel,ID,GradeID,GradeFlag,EssayNum From [User] where ID in("&ID&")")
	Select case int(Point)
	Case 1
		S="对用户："
		Do while not Rs.eof
		S=S&"“"&Rs(0)&"”"
		Rs.movenext
		Loop
		S=S&"作删除标记!"
		BBS.Execute("update [User] Set IsDel=1 where ID in("&ID&")")
	Case 2
		S="完全删除用户："
		Do while not Rs.eof
			AllTable=Split(BBS.BBStable(0),",")
			For i=0 To uBound(AllTable)
				BBS.Execute("Delete From[Bbs"&AllTable(i)&"] where Name='"&Rs(0)&"'")
				BBS.Execute("Delete From[Bbs"&AllTable(i)&"] where ReplyTopicID in (Select TopicID From[Topic] where Name='"&Rs(0)&"')")	
			Next
			BBS.Execute("Delete From[Topic] where  Name='"&Rs(0)&"'")
			BBS.Execute("Delete From[Sms] where  MyName='"&Rs(0)&"'")
			BBS.Execute("Delete From[admin] where Name='"&Rs(0)&"'")
			S=S&"“"&Rs(0)&"”"
		Rs.movenext
		Loop
		BBS.Execute("Delete * From [User] where ID in("&ID&")")
	Case 3
		S="批量删除用户："
		Do while not Rs.eof
			AllTable=Split(BBS.BBStable(0),",")
			For i=0 To uBound(AllTable)
				BBS.Execute("Delete From[Bbs"&AllTable(i)&"] where Name='"&Rs(0)&"'")
				BBS.Execute("Delete From[Bbs"&AllTable(i)&"] where ReplyTopicID in (Select TopicID From[Topic] where Name='"&Rs(0)&"')")	
			Next
			BBS.Execute("Delete From[Topic] where Name='"&Rs(0)&"'")
			S=S&"“"&Rs(0)&"”"
		Rs.movenext
		Loop
		S=S&"的所有帖子成功!"			
	Case 4
		S="屏蔽用户："
		Do while not Rs.eof
		S=S&"“"&Rs(0)&"”"
		Rs.movenext
		Loop
		S=S&"的所有帖子成功!"
		BBS.Execute("update [User] Set IsShow=1 where ID in("&ID&")")
	Case 5
		S="屏蔽用户："
		Do while not Rs.eof
		S=S&"“"&Rs(0)&"”"
		Rs.movenext
		Loop
		S=S&"的个人签名成功!"
		BBS.Execute("update [User] Set IsSign=1 where ID in("&ID&")")
	Case 6
		S="提升用户："
		Do while not Rs.eof
		S=S&"“"&Rs(0)&"”"
		If Rs(5)=0 Then'如果是普通等级组标记
			BBS.UpdateGrade Rs(3),0,4
		End If
		Rs.movenext
		Loop
		S=S&"为VIP会员!"
		BBS.Execute("update [User] Set IsVip=1 where ID in("&ID&")")	
	Case 7
		Dim EssayNum,GoodNum,Grade,Rs1
		S="整理修复用户："
		AllTable=Split(BBS.BBStable(0),",")
		Do while not rs.eof
		S=S&"“"&Rs(0)&"”"
		EssayNum=0
		GoodNum=0
		For i=0 To uBound(AllTable)
			EssayNum=EssayNum+BBS.execute("select Count(BbsID) from [bbs"&AllTable(i)&"] where Name='"&Rs(0)&"'")(0)
		Next
			GoodNum=BBS.execute("select count(TopicID) from [Topic] where Name='"&Rs(0)&"' And IsGood=1")(0)
			If Rs(5)=0 or BBS.Execute("Select ID From [grade] where ID="&Rs(4)).Eof Then
			BBS.UpdateGrade Rs(3),EssayNum,0
			End If
			BBS.execute("update [User] set EssayNum="&EssayNum&",GoodNum="&GoodNum&" where Name='"&Rs(0)&"'")
		Rs.Movenext
		Loop
		S=S&"各项数据成功!"
	Case 8
		Dim GradeID
		GradeID=Request.form("GradeID")
		If GradeID="" Then Goback"",""
		S="提升用户："
		Do while not Rs.eof
		If Rs(5)<4 Then'对版主以上级别无效
			S=S&"“"&Rs(0)&"”"
			IsBe=True
			BBS.execute("update [User] set GradeID="&GradeID&",GradeFlag=1 where Name='"&Rs(0)&"'")
			BBS.UpdageOnline Rs(0),3
		End If
		Rs.movenext
		Loop
		If Not IsBe Then Goback"","选定的用户已经是版主以上的等级了":Exit Sub
		S=S&"为特别等级组 "&BBS.GetGradeName(GradeID,0)&" 成功!"
	Case 9
		S="把用户："
		Do while not Rs.eof
		If Rs(5)=1 Then
			IsBe=True
			S=S&"“"&Rs(0)&"”"
			If Rs(1)=1 Then
				BBS.UpdateGrade Rs(3),0,4
			Else
				BBS.UpdateGrade Rs(3),Rs(6),0
			End If
			BBS.UpdageOnline Rs(0),3
		End If
		Rs.movenext
		Loop
		If Not IsBe Then Goback"","选定的用户不属于特别等级组":Exit Sub
		S=S&"降回正常发帖等级组成功!"
	Case 10	
		S="通过注册用户："
		Do while not Rs.eof
		S=S&"“"&Rs(0)&"”"
		Rs.movenext
		Loop
		S=S&"的审核!"
		BBS.Execute("update [User] Set IsDel=0 where ID in("&ID&")")
	Case 12	
		S="对删除用户："
		Do while not Rs.eof
		S=S&"“"&Rs(0)&"”"
		Rs.movenext
		Loop
		S=S&"恢复成功!"
		BBS.Execute("update [User] Set IsDel=0 where ID in("&ID&")")
	Case 13
		S="对用户："
		Do while not Rs.eof
		S=S&"“"&Rs(0)&"”"
		Rs.movenext
		Loop
		S=S&"的所有屏蔽帖子恢复成功!"
		BBS.Execute("update [User] Set IsShow=0 where ID in("&ID&")")
	Case 14
		S="恢复显示用户："
		Do while not Rs.eof
		S=S&"“"&Rs(0)&"”"
		Rs.movenext
		Loop
		S=S&"的个人签名成功!"
		BBS.Execute("update [User] Set IsSign=0 where ID in("&ID&")")
	Case 11
		S="取消VIP会员用户："
		Do while not Rs.eof
		S=S&"“"&Rs(0)&"”"
		If Rs(5)=4 Then'如果是VIP等级标记
		BBS.UpdateGrade Rs(3),EssayNum,0
		End If
		Rs.movenext
		Loop
		S=S&"!"
		BBS.Execute("update [User] Set IsVip=0 where ID in("&ID&")")	
	End Select
	Rs.Close
	BBS.NetLog"操作后台_"&S
	Suc "",S,"admin_actionList.asp?action=UserList"
End Sub

Sub SaveMenu 
	Dim MenuName,MenuUrl,Show,ParenID,Target,ID,Flag,S
	MenuName=BBS.Fun.GetForm("MenuName")
	MenuUrl=BBS.Fun.GetStr("MenuUrl")
	ID=Request("ID")
	Show=Request("Show")
	Target=Request("Target")
	Flag=Request("Flag")
	ParenID=Request("ParenID")
	If MenuName="" Then GoBack"","":Exit Sub
	IF ID<>"" Then
		Dim Subs
		Subs=BBS.Execute("Select Count(*) From [Menu] where parenID="&ID)(0)
		If Subs>0 And Int(ParenID)>0 Then GoBack"","该菜单包含下拉菜单项目，不能作为下拉菜单项。":Exit Sub
		If Flag<>"" Then
			If Flag="8" Then
				BBS.Execute("Update [Menu] Set name='"&MenuName&"',Show="&Show&" where ID="&ID)
			Else
				BBS.Execute("Update [Menu] Set name='"&MenuName&"',Show="&Show&",Target="&Target&",ParenID="&ParenID&" where ID="&ID)
			End IF
		Else
			BBS.Execute("Update [Menu] Set name='"&MenuName&"',url='"&MenuUrl&"',Show="&Show&",Target="&Target&",ParenID="&ParenID&" where ID="&ID)
		End if
		S="修改菜单："&MenuName&" 成功!"
		BBS.NetLog"操作后台_"&S
		Suc"",S,"admin_action.asp?action=Menu"
	Else
		Dim Orders
		Orders=BBS.Execute("Select Count(*) from [Menu] where ParenID="&ParenID)(0)+1
		BBS.Execute("insert into [Menu](name,url,Target,Show,ParenID,Orders) values('"&MenuName&"','"&MenuUrl&"',"&Target&","&Show&","&ParenID&","&Orders&")")
		S="添加菜单："&MenuName&" 成功!"
		BBS.NetLog"操作后台_"&S
		Suc"",S,"admin_action.asp?action=Menu"
	End IF
End Sub

Sub MenuOrder
	Dim Orders,ID,I,S
	For i=1 to request.form("ID").count
		ID = request.form("ID")(i)
		Orders = request.form("Orders")(i)
		If IsNumeric(Orders) And isnumeric(ID) Then
			BBS.Execute("Update [Menu] Set Orders="&Orders&" where ID="&ID)
		End If
	Next
	S="菜单排序更新成功!"
	BBS.NetLog"操作后台_"&S
	Suc "",S,"admin_action.asp?action=Menu"
End Sub

Sub DelMenu
	Dim ID,S
	ID=Request.QueryString("ID")
	BBS.Execute("Delete From[Menu] where Flag=0 And ID="&ID)
	S="删除菜单成功"
	BBS.NetLog"操作后台_"&S
	Suc"",S,"admin_action.asp?action=Menu"
End Sub


'==--->>>生成JS文件
Sub SetJsMenu
	Dim objFSO,objName
	Dim UserMenu,TouristMenu,uM,tM
	Dim Target,Target1,S
	Dim Rs,Rs1,I,II
	Dim Board_Rs,Po,BoardMenu,BoardSelect
	'生成顶部菜单
	Set Rs=BBS.Execute("Select ID,Name,Url,show,orders,flag,Target From [Menu] where show<3 and parenID=0 order by orders")
	Do while not Rs.eof
	 '风格指定为8
	 If Rs(5)=8 Then
			S="<div class=menuitems><A href=cookies.asp?action=style&skinid=0>默认风格<\/A><\/div>"
			Set Rs1=BBS.Execute("Select skinid,SkinName From[Skins] where pass=1")
			do while Not Rs1.Eof
				S=S&"<div class=menuitems><A href=cookies.asp?action=style&skinid="&Rs1(0)&">"&BBS.Fun.GetJsStr(Rs1(1))&"<\/A><\/div>"
			rs1.movenext
			Loop
			Rs1.close
			S="<DIV id=M"&Rs(0)&" class=menu>"&S&"<\/DIV>　<a href='#' onmouseover=\""dropdownmenu(this, event, \'M"&Rs(0)&"\');\"" >"&BBS.Fun.GetJsStr(Rs(1))&"<\/a>"
			If Rs(3)<>2 Then UserMenu=UserMenu&S
			If Rs(3)<>1 Then TouristMenu=TouristMenu&S
	 Else
		If Rs(6)=0 then
			Target=""
		Else
			Target=" target=_bank"
		End If
		Set Rs1=BBS.Execute("Select ID,Name,Url,show,orders,flag,Target From [Menu] where show<3 and parenID="&Rs(0)&" order by orders")
		If Rs1.eof Then
			If IsNull(Rs(2)) or Rs(2)="" Then
				If Rs(3)<>2 Then UserMenu=UserMenu&"　<a>"&BBS.Fun.GetJsStr(Rs(1))&"<\/a>"
				If Rs(3)<>1 Then TouristMenu=TouristMenu&"　<a>"&BBS.Fun.GetJsStr(Rs(1))&"<\/a>"
			Else
				If Rs(3)<>2 Then UserMenu=UserMenu&"　<a href="&BBS.Fun.GetJsStr(Rs(2))&""&Target&">"&BBS.Fun.GetJsStr(Rs(1))&"<\/a>"
				If Rs(3)<>1 Then TouristMenu=TouristMenu&"　<a href="&BBS.Fun.GetJsStr(Rs(2))&""&Target&">"&BBS.Fun.GetJsStr(Rs(1))&"<\/a>"
			End IF
		Else
			uM=""
			tM=""
			Do while not Rs1.eof
				If Rs1(6)=0 then
					Target1=""
				Else
					Target1=" target=_bank"
				End If
				S="<div class=\""menuitems\""><a href="&BBS.Fun.GetJsStr(Rs1(2))&""&Target1&">"&BBS.Fun.GetJsStr(Rs1(1))&"<\/a><\/div>"
				If Rs1(3)<>2 Then uM=uM&S
				If Rs1(3)<>1 Then tM=tM&S
			Rs1.movenext
			Loop
			uM="<DIV id=\""M"&Rs(0)&"\"" class=\""menu\"">"&uM&"<\/div>"
			tM="<DIV id=\""M"&Rs(0)&"\"" class=\""menu\"">"&tM&"<\/div>"
			UserMenu=UserMenu&uM
			TouristMenu=TouristMenu&tM
			If Rs(2)>"" Then
				S="　<a onmouseover=\""dropdownmenu(this, event, \'M"&Rs(0)&"\');\"" href="&BBS.Fun.GetJsStr(Rs(2))&""&Target&">"&BBS.Fun.GetJsStr(Rs(1))&"<\/a>"
			Else
				S="　<a href='#' onmouseover=\""dropdownmenu(this, event, \'M"&Rs(0)&"\');\"">"&BBS.Fun.GetJsStr(Rs(1))&"<\/a>"
			End IF
				If Rs(3)<>2 Then UserMenu=UserMenu&S
				If Rs(3)<>1 Then TouristMenu=TouristMenu&S
		End IF
		Rs1.close
		Set Rs1=nothing
	 End If
	Rs.movenext
	Loop
	Rs.Close
	UserMenu=Replace(UserMenu,"{用户名}","""+user+""")
'生成版块菜单
	Set Rs=BBS.Execute("Select Depth,boardid,ParentID,Boardname,BoardImg,Introduce,Boardadmin,PassUser,Child,ParentStr,RootID,Strings From[Board] order by RootID,Orders")
	If not Rs.Eof Then
		Board_Rs=Rs.GetRows(-1)
	End If
	If IsArrAy(Board_Rs) Then 
		For i=0 To Ubound(Board_Rs,2)
		Po=""
		If Board_Rs(0,i)=0 Then'类
			BoardMenu=BoardMenu&"<div class=\""menuitems\""><a href=\""board.asp?boardid="&Board_Rs(1,i)&"\""><b>"&BBS.Fun.GetJsStr(Board_Rs(3,i))&"</b></a></div>"
			BoardSelect=BoardSelect&"<option><b>"&Board_Rs(3,i)&"</b></option>"
		Else
			For II=2 to Board_Rs(0,i)
			Po=Po&"∣"
			Next
			BoardMenu=BoardMenu&"<div class=\""menuitems\""><A href=\""board.asp?boardid="&Board_Rs(1,i)&"\"">"&po&"&nbsp;&nbsp;├ "&BBS.Fun.GetJsStr(Board_Rs(3,i))&"</a></div>"
			BoardSelect=BoardSelect&"<option value=\"""&Board_Rs(1,i)&"\"">"&po&"&nbsp;&nbsp;├ "&Board_Rs(3,i)&"</option>"
		End IF
		Next
		BoardSelect="<select onchange=if(this.options[this.selectedIndex].value!=''){location='board.asp?boardid='+this.options[this.selectedIndex].value;}><option selected>跳转论坛至...</option>"&BoardSelect&"</select>"
		BoardMenu="<div id=\""Board\"" class=\""menu\"">"&BoardMenu&"</div>"
	End If
	
	On Error Resume Next
	Set objFSO = Server.CreateObject("Scr"&"ipting.Fil"&"eSystemOb"&"ject")
	Set objName=objFSO.CreateTextFile(Server.MapPath("Inc/setmenu.js"),True)
	objName.Write"function UserMenu(user){"&vbcrlf&"document.write("""&UserMenu&""");"&vbcrlf&"}"&vbcrlf&"function TouristMenu(){"&vbcrlf&"document.write("""&TouristMenu&""");"&vbcrlf&"}"&vbcrlf&"function BoardListMenu(){"&vbcrlf&"document.write("""&BoardMenu&""");"&vbcrlf&"}"&vbcrlf&"function BoardSelect(){"&vbcrlf&"document.write("""&BoardSelect&""");"&vbcrlf&"}"
	objName.Close
	Set objFso=Nothing
	If Err Then
		Goback"","更新失败，空间不支持FOS文件读写！。"
		err.Clear
		Exit Sub
	End If
	Suc "","成功的更新了前台的菜单:顶部各项菜单、版块下拉菜单!","javascript:history.go(-1)"
	BBS.Netlog "操作后台_生成前台Js菜单成功!"
end sub

Sub saveagreement()
	On Error Resume Next
	dim objFSO,objname,S
	Set objFSO = Server.CreateObject("Scr"&"ipting"&".Fil"&"eSy"&"stemOb"&"ject")
	Set objname=objFSO.CreateTextFile(Server.MapPath("inc/agreement.html"),True)
	objname.Write replace(request("content"),"'","''")
	objname.close
	Set objfso=nothing
	If Err Then
	Goback"","修改失败，空间不支持FOS文件读写！请直接编辑inc/agreement.html这个文件。"
	err.Clear
	Exit Sub
	End If
	S="修改论坛注册协议成功!"
	BBS.Netlog "操作后台_"&S
	Suc "修改成功",S,"admin_SetHtmlEdit.asp?action=Agreement"
End Sub

Sub DelPlacard
	Dim ID,S
	ID=Request.QueryString("ID")
	BBS.execute("Delete From [Placard] where ID="&ID&"")
	BBS.Cache.clean("Placard")
	S="删除公告成功!"
	BBS.NetLog"操作后台_"&S
	Suc"",S,"admin_actionList.asp?action=Placard"
End Sub

Sub Placard
	With BBS
	Dim ID,Caption,Content,Hits,Addtime,Name,B_ID,S
	ID=Request.Form("ID")
	Caption=.Fun.GetForm("Caption")
	Content=.Fun.GetForm("Content")
	Hits=.Fun.GetForm("Hits")
	AddTime=.Fun.GetForm("AddTime")
	Name=.Fun.GetForm("Name")
	B_ID=.Fun.GetForm("boardid")
	If Caption="" or Content="" Then GoBack"","":Exit Sub
	S=.Fun.UbbString(Content)
	If ID<>"" Then
		BBS.execute("update [Placard] Set Caption='"&Caption&"',[Name]='"&Name&"',Content='"&Content&"',AddTime='"&AddTime&"',ubbString='"&S&"',boardid="&B_ID&",Hits="&Hits&" where ID="&ID&"")
		S="公告编辑成功"
	Else
		BBS.execute("insert into[Placard](Caption,Content,AddTime,[Name],boardid,Hits,UbbString)values('"&Caption&"','"&Content&"','"&AddTime&"','"&Name&"',"&B_ID&","&Hits&",'"&S&"')")
		S="公告发布成功"
	End If
	.NetLog"操作后台_"&S
	.Cache.clean("Placard")
	Suc"",S&"！","admin_actionList.asp?action=Placard"
	End With
End Sub

Sub DelLockIP
	Dim ID
	ID=Int(Request("ID"))
	BBS.Execute("Delete From [LockIP] Where ID="&ID&"")
	BBS.Cache.clean("IPData")
	BBS.NetLog"操作后台_删除封IP记录成功!"
	Response.redirect "admin_action.asp?action=LockIp"
End Sub

Sub LockIp
	Dim ID,StartIP,EndIP,Readme,S
	ID=Request("ID")
	StartIp=BBS.Fun.GetForm("StartIp")
	EndIp=BBS.Fun.GetForm("EndIp")
	Readme=BBS.Fun.GetForm("Readme")
	If StartIp="" Then
		GoBack"",""
		Exit Sub
	End If
	if EndIp="" then EndIp=StartIp
	If ID<>"" Then
		BBS.execute("update [LockIp]Set StartIp="&BBS.Fun.IpEnCode(StartIp)&",EndIp="&BBS.Fun.IpEnCode(EndIp)&",Readme='"&Readme&"' where ID="&ID&"")
		S="修改封锁IP成功!"
		Suc"",S,"admin_action.asp?action=LockIp"
	Else
		BBS.execute("Insert into [LockIp](StartIp,EndIp,Readme,lock)values("&BBS.Fun.IpEnCode(StartIp)&","&BBS.Fun.IpEnCode(EndIp)&",'"&Readme&"',1)")
		S="网段封锁成功!"
		Suc"网段封锁成功","倒霉的孩子的IP已经被封!","admin_action.asp?action=LockIp"
	End If
	BBS.NetLog"操作后台_"&S
	BBS.Cache.clean("IPData")
End Sub

Sub IsLockIp
	Dim ID,IsLock,S
	ID=Int(Request("ID"))
	IsLock=BBS.Execute("Select Lock From[LockIp] where Id="&ID&"")(0)
	If IsLock=1 Then
		S="解封IP成功!"
		BBS.Execute("update [LockIp] set Lock=0 where Id="&ID&"") 
	Else
		S="封锁IP成功!"
		BBS.Execute("update [LockIp] set Lock=1 where Id="&ID&"") 
	End IF
	BBS.NetLog"操作后台_"&S
	BBS.Cache.clean("IPData")
	Response.redirect "admin_action.asp?action=LockIp"
End Sub

Sub AuteSqlTable
	Dim Aute,S,AllTable,i
	Aute=BBS.Fun.GetStr("Aute")
	AllTable=Split(BBS.BBStable(0),",")
	S=""
	For i=0 To uBound(AllTable)
		If Aute=AllTable(i) Then S="yes"
	Next
	If S="" Then Goback"系统出错","无效的数据表名称！":Exit Sub
	IF Int(Aute)<>Int(BBS.BBStable(1)) Then
		S=BBS.BBStable(0)&"|"&Int(Aute)
		BBS.execute("Update [Config] Set BbStable='"&S&"' ")
	End If
	S="更改论坛默认数据表为 bbs"&Aute&" 成功!"
	BBS.NetLog"操作后台_"&S
	Suc"",S,"admin_action.asp?action=SqlTable"
	BBS.Cache.clean("parameter")
End Sub

Sub AddSqlTable
	Dim TableName,AllTable,I,S
	TableName=BBS.Fun.GetStr("TableName")
	If not BBS.Fun.isInteger(TableName) then
		GoBack"","请用正整数的数字填写！"
		Exit Sub
	End If
	If Int(TableName)=0 Then
		GoBack"","数据表名不能为0"
		Exit Sub
	End If
	AllTable=Split(BBS.BBStable(0),",")
	For i=0 To uBound(AllTable)
	If Int(TableName)=Int(AllTable(i)) then
		GoBack"","数据表名已经存在！"
		Exit Sub
	End if
	Next
	S=BBS.BBStable(0)&","&TableName&"|"&BBS.BBStable(1)
	BBS.execute("update [config] Set BbStable='"&S&"'")
	BBS.execute("CREATE TABLE [bbs"&TableName&"](BbsID int IDENTITY (1, 1) NOT NULL CONSTRAINT PrimaryKey PRIMARY KEY,TopicID int Default 0,ReplyTopicID int Default 0,boardid int Default 0,Name varchar(20),Caption varchar(255),Content text,Face int Default 0,AddTime Datetime,LastTime datetime,IsDel byte Default 0,Ip varchar(40),IsAppraise byte Default 0,UbbString varchar(255))")
	BBS.execute("create index TopicID on [bbs"&TableName&"] (TopicID)")
	BBS.execute("create index boardid on [bbs"&TableName&"] (boardid)")
	BBS.execute("create index ReplyTopicID on [bbs"&TableName&"] (ReplyTopicID)")
	S="添加 Bbs"&TableName&" 数据表成功!"
	BBS.NetLog"操作后台_"&S
	Suc"",S,"admin_action.asp?action=SqlTable"
	BBS.Cache.clean("parameter")
End Sub

Sub DelSqlTable
	Dim ID,S,AllTable,I
	ID=request.querystring("ID")
	If int(ID)=int(BBS.BBStable(1)) Then
		GoBack "","该表被设定为默认使用表，不能删除！"
		Exit Sub
	End if
	AllTable=Split(BBS.BBStable(0),",")
	S=""
	For i=0 To uBound(AllTable)
		If int(ID)=Int(AllTable(i)) Then S="yes"
	Next
	If S="" Then
		Goback"系统出错","无效的数据表名称！":Exit Sub
	End If
	S=""
	For i=0 To uBound(AllTable)
		If Int(ID)<>int(AllTable(i)) then
			S=S&AllTable(i)&","
		End if
	Next
	S=Left(S,len(S)-1)
	S=S&"|"&BBS.BBStable(1)
	BBS.execute("update [Config] Set BbStable='"&S&"'")
	S=BBS.execute("Select Count(*) From[bbs"&ID&"]")(0)
	BBS.Execute("Drop table [bbs"&ID&"]")
	BBS.Execute("Delete * From [Topic] where SqlTableID="&ID&"")
	BBS.Cache.clean("parameter")
	S="删除名称为 Bbs"&ID&" 的数据表("&S&"篇帖子)!"
	BBS.NetLog"操作后台_"&S
	Suc"","成功的"&S,"admin_action.asp?action=SqlTable"
End Sub

Sub SqlTableUnite
	Dim ID1,ID2,S,AllTable,i
	ID1=request.form("SqlTableID1")
	ID2=request.form("SqlTableID2")
	If ID1="0" or ID1="" or ID2="" or ID2="0" Then
	GoBack"","没有选定！"
	Exit Sub
	End If
	If ID1=ID2 Then
		GoBack "","同一个数据表还用合并吗？晕~！"
		Exit Sub
	End If
	
	If int(ID1)=int(BBS.BBStable(1)) Then
		GoBack "","指定数据表是默认使用表，不能合并到目标表！"
		Exit Sub
	End if
	'检验
	AllTable=Split(BBS.BBStable(0),",")
	For i=0 To uBound(AllTable)
		If int(ID1)=Int(AllTable(i)) Then S=S&"Y"
		If Int(ID2)=Int(AllTable(i)) Then S=S&"Y"
	Next
	If S<>"YY" Then Goback"系统出错","无效的数据表名称！":Exit Sub
	Set Rs=BBS.Execute("select * From [bbs]"&ID1&"")
	do while not rs.eof
		BBS.execute("insert into [bbs"&ID2&"](TopicID,ReplyTopicID,boardid,[Name],Caption,Content,Face,AddTime,LastTime,IsDel,Ip,IsAppraise,UbbString)values("&Rs("TopicID")&","&Rs("ReplyTopicID")&","&Rs("boardid")&",'"&Rs("Name")&"','"&Rs("Caption")&"','"&Rs("Content")&"',"&Rs("Face")&",'"&Rs("AddTime")&"','"&Rs("LastTime")&"',"&Rs("IsDel")&",'"&Rs("Ip")&"',"&Rs("IsAppraise")&",'"&Rs("UbbString")&"')")
	Rs.movenext
	Loop
	Rs.close
	BBS.Execute("update [Topic] set SqltableID="&ID2&" where SqlTableID="&ID1)
	BBS.Execute("Drop table [bbs"&ID1&"]")
	S=""
	For i=0 To uBound(AllTable)
		If Int(ID1)<>int(AllTable(i)) then
			S=S&AllTable(i)&","
		End if
	Next
	S=Left(S,len(S)-1)
	S=S&"|"&BBS.BBStable(1)
	BBS.execute("update [Config] Set BbStable='"&S&"'")
	S="数据表BBS"&ID1&"合并到数据表BBS"&ID2&"成功!"
	BBS.NetLog "操作后台_"&S
	Suc"",S,"admin_action.asp?action=SqlTable"
	BBS.Cache.clean("parameter")
End Sub

Sub UpdateBbsDate
	Dim EssayNum,TopicNum,NewUser,TodayNum,UserNum,AllTable,I
	UserNum=BBS.Execute("Select Count(ID) From[User]")(0)
	NewUser=BBS.execute("select Top 1 Name from [User] order by ID desc")(0)
	AllTable=Split(BBS.BBStable(0),",")
	For i=0 To uBound(AllTable)
		EssayNum=EssayNum+BBS.Execute("Select Count(*) From[Bbs"&AllTable(i)&"] where IsDel=0")(0)
		TodayNum=TodayNum+BBS.Execute("Select Count(*) From[Bbs"&AllTable(i)&"] where IsDel=0 And DATEDIFF('d',LastTime,'"&BBS.NowBbsTime&"')<1")(0)
	Next
	TopicNum=BBS.Execute("Select Count(TopicID) From[Topic]")(0)
	BBS.Execute("update [Config] Set UserNum="&UserNum&",AllEssayNum="&EssayNum&",TopicNum="&TopicNum&",TodayNum="&TodayNum&",NewUser='"&NewUser&"'")
	BBS.NetLog"操作后台_论坛系统整理成功!"
	Suc "","论坛系统整理成功!整理后：<li>总帖数："&EssayNum&" | 主题数："&TopicNum&" | 今日帖数："&TodayNum&" | 注册用户数："&UserNum&" | 最新注册用户："&NewUser&"","admin_action.asp?action=UpdateBbs"
	BBS.Cache.clean("InfoUpdate")
End Sub

Sub UpdateTopic
	Dim Caption,Content,ID1,ID2,LastReply,Go,ReplyNum,Rs1,AllTable,i,S,MaxID,II
	ID1=BBS.Fun.GetStr("id1")
	ID2=BBS.Fun.GetStr("id2")
	MaxID=BBS.execute("select max(TopicID)from [Topic]")(0)
	IF ID1="" Then
		ID1=1:ID2=100
		Go= "开始整理"
	Else
		IF not isnumeric(ID1) or not isnumeric(ID2) then GoBack"","<li>请用数字填写！":Exit Sub
		Set Rs=BBS.execute("Select TopicId,SqlTableID,Name From [Topic] where TopicID>="&ID1&" And Topicid<="&ID2&"")
		IF Not Rs.eof then
			AllTable=Split(BBS.BBStable(0),",")
			Do while not rs.eof
			For i=0 To uBound(AllTable)
				LastReply=Rs(2)&"|暂无回复"
				II=0			
				IF Int(Rs(1))=Int(AllTable(i)) Then
					ReplyNum=BBS.execute("select count(BbsID) from [bbs"&Rs(1)&"] where IsDel=0 and ReplyTopicID="&Rs(0)&"")(0)
					Set Rs1=BBS.Execute("Select Name,Content From [bbs"&Rs(1)&"] where IsDel=0 and ReplyTopicID="&Rs(0)&" order by BbsID desc")
					Do while Not Rs1.Eof
					II=II+1
					If II=1 Then LastRePly=Rs1(0)&"|"&Replace(BBS.Fun.StrLeft(BBS.Fun.FixReply(Rs1(1)),40),"'","")
					Rs1.movenext
					Loop
					Rs1.close
					BBS.execute("update [Topic] Set ReplyNum="&ReplyNum&",LastReply='"&LastReply&"' where TopicID="&Rs(0))
					Exit For
				End If
			Next
			Rs.Movenext
			Loop
			Rs.Close
			Set Rs1=Nothing
		End If
	S=ID1
	ID1=int(ID2)+1
	ID2=int(ID2)+int(ID2)-int(S)+1
	Go="继续整理"
	End If
	If Int(ID1)>Int(MaxID) Then
		Suc"整理结束","全部整理成功!","admin_action.asp?action=UpdateBbs"
		BBS.NetLog"操作后台_整理修复主题帖子"
		Exit Sub
	End If
	Caption="主题帖子整理"
	Content="<form method='POST' style='margin:0' action='?action=UpdateTopic' onSubmit='ok.disabled=true;ok.value=""正在整理-请稍等。。。""'>请填写你要整理的主题帖子的开始ID和结束ID：（两者之间不要相差太大）<br>你的论坛主题帖子最大的 ID 为："&MaxID&"<br>初始ID：<input type='text' name='ID1' size='20' value='"&ID1&"' class='text'><BR>结束ID：<input type='text' name='ID2' size='20' value='"&ID2&"' class='text' /><BR><input type='submit' name='ok' value='"&Go&"' class='button' /><input type='reset' value='重 置' class='button'></form>"
	ShowTable caption,Content
End Sub


Sub UpdateAllUser
	Dim Content,ID1,ID2,GoodNum,EssayNum,Rs1,Go,AllTable,I,S,MaxID,Flag
	ID1=BBS.Fun.GetStr("id1"):ID2=BBS.Fun.GetStr("id2")
	MaxID=BBS.execute("select max(id) from [User]")(0)
	IF ID1="" Then
		ID1=1:ID2=100
		Go= "开始整理"
	Else
		IF not isnumeric(ID1) or not isnumeric(ID2) then GoBack"","<li>请用数字填写！":Exit Sub
		Set Rs=BBS.execute("Select ID,name,IsVIP,GradeFlag,GradeID From [User] where Id>="&ID1&" and id<="&ID2&"")
		IF Not Rs.eof then
			AllTable=Split(BBS.BBStable(0),",")
			Do while not rs.eof
			EssayNum=0
			GoodNum=0
			For i=0 To uBound(AllTable)
				EssayNum=EssayNum+BBS.execute("select Count(BbsID) from [bbs"&AllTable(i)&"] where Name='"&Rs(1)&"'")(0)
			Next
				GoodNum=BBS.execute("select count(TopicID) from [Topic] where Name='"&Rs(1)&"' And IsGood=1")(0)
	'更新等级
			Flag=Rs(3)
			If Flag>3 Then
				Set Rs1=BBS.Execute("select boardid from [admin] where name='"&Rs(1)&"' order by boardid")
				IF Not Rs1.eof Then
					IF Rs1(0)=0 Then
						Flag=9 
					ElseIF Rs1(0)=-1 Then
						Flag=8
					Else
						Flag=7
					End if
				End If
				Rs1.Close		
				If Rs(2)=1 and Flag=0 Then Flag=4	
			ElseIf Flag=1 Then'如果为特殊组
				If BBS.Execute("Select ID From [grade] where ID="&Rs(4)).Eof Then Flag=0
			End IF
			BBS.UpdateGrade Rs(0),EssayNum,Flag	
			
			BBS.execute("update [User] set EssayNum="&EssayNum&",GoodNum="&GoodNum&" where ID="&Rs(0)&"")
		Rs.Movenext
		Loop
		rs.close
		Set Rs1=nothing
	End IF
	S=ID1
	ID1=int(ID2)+1
	ID2=int(ID2)+int(ID2)-int(S)+1
	Go="继续整理"
	End If
	If Int(ID1)>Int(MaxID) Then
		BBS.NetLog"操作后台_整理修复用户数据"
		Suc"整理结束","全部整理成功!","admin_action.asp?action=UpdateBbs"
		Exit Sub
	End If
	Content="<form method='POST' style='margin:0' action='?action=UpdateAllUser' onSubmit='ok.disabled=true;ok.value=""正在整理-请稍等。。。""'>请填写你要整理用户的开始ID和结束ID：（两者之间不要相差太大）<br />论坛注册用户最大的 ID 为："&MaxID&"<br />初始ID：<input type='text' class='text' name='id1' size=20 value='"&ID1&"' /><br />结束ID：<input type='text' class='text' name='id2' size='20' value='"&ID2&"' /><br /><input name='ok' class='button' type=submit value="&Go&" /><input type='reset' value='重 置' class='button' /></form>"
	ShowTable "用户整理修复",Content
End Sub

Sub DelWuiong
	Dim i,AllTable,content
	Response.Write"<div class='mian'><div class='top'>论坛垃圾清理</div>"&_
	"<div class='divth'><b><span id='BBST'></span></b><div class='mian' style='margin:2px auto 0;width:400px;height:9'><img src='Images/icon/hr1.gif' width=0 height=16 id='BBSimg' align='absmiddle' alt='进度条' /></div>"&_
	"<div><span id='BBStxt' style='font-size:9pt'>0</span>%</div></div></div>"
	Response.Flush
	'BBS.execute("delete * from [admin] where (boardid<>0 and boardid<>-1) and (boardid not in(select boardid from [Board] where parentID<>0) or name not in(select name From [user] where isdel=1))")
	Call PicPro(0,8,"正在清理无效版主！请稍等。。。")	
		Set Rs=BBS.execute("Select name,boardid from [admin] where boardid<>0 and boardid<>-1")
		do while not Rs.eof
		If BBS.Execute("Select * From [Board] where ParentID<>0 and boardid="&Rs(1)&"").eof Then
			BBS.execute("delete * from [admin] where name='"&Rs(0)&"' and boardid>0 ")
		ElseIf BBS.Execute("Select * From [User] where Name='"&Rs(0)&"' and Isdel=0").eof Then
			BBS.execute("delete * from [admin] where name='"&Rs(0)&"' and boardid>0")
		End If
		Rs.Movenext
		Loop
		Rs.Close
		Show"清理无效版主完毕！"
			
	Call PicPro(1,8,"正在清理无效主题！请稍等。。。")	
		AllTable=Split(BBS.BBStable(0),",")
		For i=0 To uBound(AllTable)
			BBS.execute("delete * from [bbs"&AllTable(i)&"] where TopicID<>0 and not exists (select name from [topic] where [bbs"&AllTable(i)&"].TopicId=[Topic].TopicID)")
			BBS.execute("delete * from [Topic] where SqlTableID="&AllTable(i)&" and not exists (select name from [bbs"&AllTable(i)&"] where [Topic].TopicID=[bbs"&AllTable(i)&"].TopicId)")
		Next
		Show"无效主题清理完毕！"
	
	Call PicPro(2,8,"正在清理无效的评帖记录")
		BBS.execute("delete * from [Appraise] where  not exists (select name from [Topic] where [Appraise].TopicID=[Topic].TopicId)")
		Show "无效评帖记录清理完毕！"	

	Call PicPro(3,8,"正在清理无效投票！请稍等。。。")
		BBS.execute("delete * from [TopicVote] where  not exists (select name from [Topic] where [TopicVote].TopicID=[Topic].TopicId)")
		BBS.execute("delete * from [TopicVoteUser] where  not exists (select name from [Topic] where [TopicVoteUser].TopicID=[Topic].TopicId)")
		Show"无效投票清理完毕！"
	
	Call PicPro(4,8,"正在清理无效留言！请稍等。。。")
		BBS.execute("delete * from [Sms] where not exists (select name from [User] where [Sms].MyName=[User].Name)")
		Show"无效留言清理完毕！"
	Call PicPro(5,8,"正在清理无效公告！请稍等。。。")
		BBS.execute("delete * from [Placard] where not exists (select name from [User] where [Placard].Name=[User].Name)")
		If Not IsArray(BBS.Board_Rs) Then BBS.GetBoardCache()
		If IsArray(BBS.Board_Rs) Then
			For i=0 To Ubound(BBS.Board_Rs,2)
			'如果是版块为类
			If BBS.Board_Rs(0,i)=0 Then
				BBS.execute("delete * from [Placard] where boardid<0 or boardid="&BBS.Board_Rs(1,i))
			End If
			Next
		End If
		Show"无效公告清理完毕！"
	
	Call PicPro(6,8,"正在清理删除用户的帖子！请稍等。。。")
		For i=0 To uBound(AllTable)
		BBS.execute("delete * from [bbs"&AllTable(i)&"] where not exists (select name from [User] where [bbs"&AllTable(i)&"].Name=[User].Name)")
		Next
		BBS.execute("delete * from [Topic] where not exists (select name from [User] where [Topic].Name=[User].Name)")
		Show "无效用户的帖子清理完毕！"
	
	Call PicPro(7,8,"正在清理无效回复帖子！时间会比较长。。请稍等。。。")
	For i=0 To uBound(AllTable)	
		Set Rs=BBS.Execute("select ReplyTopicID from [bbs"&AllTable(i)&"] where ReplyTopicID<>0")
		Do While Not Rs.eof
			If BBS.execute("select TopicID from [bbs"&AllTable(i)&"] where TopicID="&Rs(0)&"").eof Then
			BBS.Execute("Delete * From [bbs"&AllTable(i)&"] where ReplyTopicID="&Rs(0))
			End IF
		Rs.MoveNext
		Loop
		Rs.Close
	Next
	Show"无效回复清理完毕！"		
	Response.Write "<script>document.getElementById(""BBSimg"").width=400;document.getElementById(""BBStxt"").innerHTML=""100"";BBST.innerHTML=""<font color=red>成功全部完成拉圾清理！</font>"";</script>"
	BBS.NetLog"操作后台_清理论坛拉圾"
End Sub
'进度条
Sub PicPro(i,sum,strtxt)
	Response.Write "<script>document.getElementById(""BBSimg"").width=" & Fix((i/sum) * 400) & ";" & VbCrLf
	Response.Write "document.getElementById(""BBStxt"").innerHTML=""" & FormatNumber(i/sum*100,2,-1) & """;" & VbCrLf
	Response.Write "document.getElementById(""BBST"").innerHTML="""& StrTxt & """;"& VbCrLf
	Response.Write "</script>" & VbCrLf
	Response.Flush
End Sub
Sub Show(Str)
	Response.Write"<div class='mian'><div class='divtr1' style='padding:5px'>"&Str&"</div></div>"
	Response.Flush
End Sub

Sub UpdateLink
	Dim ID,I,Orders,Pass,IsPic,IsIndex
	For i=1 to request.form("id").count
	ID = Replace(request.form("id")(i),"'","")
	Orders = Replace(request.form("orders")(i),"'","")
	Pass = Replace(request.form("pass"&i&""),"'","")
	IsPic= Replace(request.form("ispic"&i&""),"'","")
	IsIndex=Replace(request.form("isindex"&i&""),"'","")
	IF Not isnumeric(ID) or Not isnumeric(Orders) Then
		GoBack "","请用数字填写!"
		Exit Sub
	End IF
	If Pass<>"1" Then Pass="0"
	If Ispic<>"1" Then Ispic="0"
	If IsIndex<>"1" Then IsIndex="0"
	BBS.Execute("Update [Link] Set Orders="&Orders&",Pass="&pass&",IsPic="&IsPic&",IsIndex="&IsIndex&" where ID="&ID&"")
	Next
	 SetLinkPage
	 BBS.NetLog"操作后台_批量更新论坛连盟"
	Response.Redirect Request.ServerVariables("HTTP_REFERER")
End Sub

Sub SetLinkPage
	Dim i,ii,TempText,TempPic
	Set Rs=BBS.Execute("Select ID,Orders,BbsName,Url,pic,Readme,pass,Ispic From[Link] where pass=true order by orders")
		i=0:ii=0
	do while not rs.eof
	if len(Rs("pic"))<8 or isnull(rs("pic")) or Rs("isPic")=0 then
		i=i+1
		TempText=TempText&"<td><a target='_blank' href='"&rs("url")&"' title='"&rs("Readme")&"'>"&rs("BbsName")&"</a></td>"
		If i=7 Then i=0 : TempText=TempText&"</tr><tr>"
	ElseIF Rs("IsPic")=1 Then
		ii=ii+1
		TempPic=TempPic&"<td><a target='_blank' href='"&rs("url")&"'><img src='"&rs("pic")&"' border='0' title='"&rs("Readme")&"' width='88' height='31'></a></td>"
		If ii=7 Then ii=0 : TempPic=TempPic&"</tr><tr>"
	End if
	Rs.movenext
	loop
	Rs.close
	dim objFSO,objname
	Set objFSO = Server.CreateObject("Scr"&"ipting"&".Fil"&"eSy"&"stemOb"&"ject")
	Set objname=objFSO.CreateTextFile(Server.MapPath("linkinfo.asp"),True)
	objname.Write"<!--#include file=""inc.asp""--><"&"% BBS.Head""linkinfo.asp"","""",""本站友情链接"""&VbCrLf&"Call BBS.ShowTable(""<div>本站友情链接</div>"",""<table width='100%'  border='0' align='center' cellpadding='0' cellspacing='5'><tr>"&TempText&"</tr></table><table width='100%'  border='0' align='center' cellpadding='5' cellspacing='0'><tr>"&TempPic&"</tr></table>"")"&VbCrLf&"BBS.Footer()"&VbCrLf&"Set BBS =Nothing%"&">"
	objname.close
	Set objfso=nothing
	BBS.Cache.clean("LinkInfo")
End Sub

Sub SaveLink
	Dim BbsName,Url,Pic,Readme,admin,S,IsPic,ID,Pass
	BbsName=BBS.Fun.Getform("bbsname")
	Url=BBS.Fun.HtmlCode(Request.form("url"))
	Pic=BBS.Fun.HtmlCode(Request.form("pic"))
	Readme=BBS.Fun.GetForm("Readme")
	admin=BBS.Fun.GetForm("admin")
	Pass=Request.form("pass")
	IsPic=Request.form("ispic")
	ID=Request.form("id")
	If BbsName=""  or url=""  then
		GoBack"",""
		Exit Sub
	End if
	If ID<>"" Then
		BBS.execute("Update [Link] Set Url='"&Url&"',Pic='"&Pic&"',BbsName='"&BbsName&"',Readme='"&Readme&"',admin='"&admin&"',pass="&pass&",IsPic="&IsPic&" where ID="&ID&"")
		S="添加论坛联盟连接成功!"
	Else
		S=BBS.execute("select Count(ID) From[Link]")(0)
		S=Int(S+1)
		BBS.execute("insert into [Link] (Bbsname,Url,Pic,Readme,admin,Orders,IsPic,pass)values('"&BbsName&"','"&Url&"','"&Pic&"','"&Readme&"','"&admin&"',"&S&","&IsPic&","&Pass&")")
		S="修改论坛联盟连接成功!"
	End If
	BBS.NetLog "操作后台_"&S
	Suc"",S,"admin_actionList.asp?action=Link"
	SetLinkPage
End Sub

Sub DelLink
	Dim ID,S
		ID=request.querystring("ID")
		BBS.execute("delete from [link] where ID="&ID&"")
		SetLinkPage
		S="删除论坛联盟连接成功!"
		BBS.NetLog"操作后台_"&S
		Suc"",S,"admin_actionList.asp?action=Link"
End Sub


Sub CompressData()
	Dim DbPath,boolIs97,Caption,Content
	DbPath = request("DbPath")
	If request("DbPath")<> "" Then
		boolIs97 = request("boolIs97")
		DbPath = server.mappath(DbPath)
		CompactDB DbPath,boolIs97
	Else
	If Dbpath="" Then DbPath="data\db.mdb"
	Caption="压缩数据库"
	Content="<b>注意：</b>输入数据库所在相对路径，并且输入数据库名称（如果正在使用中数据库不能压缩，请选择备份数据库进行压缩操作）<hr size=1>"&_
	"<form style='margin:0' method='post'>压缩数据库：<input type='text' name='DbPath' value='"&DbPath&"'>&nbsp;<input type='submit' class='button' value='开始压缩' /><br><form>"&_
	"<input type='checkbox' name='boolIs97' value='True'>如果使用 Access 97 数据库请选择(默认为 Access 2000 数据库)"
	ShowTable Caption,Content
	End If
End sub

Sub CompactDB(DbPath, boolIs97)
	Set BBS =Nothing
Dim fso,Engine,strDbPath,JET_3X,Content
strDbPath = left(DbPath,instrrev(DbPath,"\"))
Set fso = CreateObject("Scr"&"ipting"&".Fil"&"eSy"&"stemOb"&"ject")
If fso.FileExists(DbPath) And IsAccess(DbPath) Then
	fso.CopyFile DbPath,strDbPath & "temp.mdb"
	Set Engine = CreateObject("JRO.JetEngine")
	If boolIs97 = "True" Then
		Engine.compactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDbPath & "temp.mdb", _
		"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDbPath & "temp1.mdb;" _
		& "Jet OLEDB:Engine Type=" & JET_3X
	Else
		Engine.compactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDbPath & "temp.mdb", _
		"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDbPath & "temp1.mdb"
	End If
fso.CopyFile strDbPath & "temp1.mdb",DbPath
fso.DeleteFile(strDbPath & "temp.mdb")
fso.DeleteFile(strDbPath & "temp1.mdb")
Set fso = nothing
Set Engine = nothing
	Response.Redirect "?action=OkCompactDB&path="&DbPath
Else
	Response.Redirect "?action=NotCompactDB"
End If
End Sub

Sub OkCompactDB
	BBS.NetLog"操作后台_压缩数据库"
	Suc "","你的数据库 " & request("Path") & "，已经压缩成功!" ,"?action=CompressData"
End Sub
Sub NotCompactDB
	BBS.NetLog"操作后台_压缩数据库失败!"
	GoBack "", "数据库名称或路径不正确，或者压缩过程发生意外！ 请重试！" 
End Sub


Sub BackupData()
Dim Caption,Content
Caption="备份论坛数据"
Content="<b>注意事项：</b><br>论坛数据库备份几乎是站长每天必做的事！<br>为保证您的数据安全，备份时请不要用默认名称来命名备份数据库。<br>发现数据丢失的时候，就可以用你最后备份的数据库恢复。<br>注意：所有路径都是相对与程序空间根目录的相对路径<hr size=1>"&_
"<form style='margin:0' method='post' action='?action=BackupData&Go=Start'>当前数据库路径(相对路径)：<input type=text size=15 name=DbPath value='data/db.mdb'><br>"&_
"备份数据库目录(相对路径)：<input type=text size='15' name='BkFolder' value='Data_Backup'>&nbsp;如目录不存在，程序将自动创建<BR>"&_
"备份数据库名称(填写名称)：<input type=text size=15 name=BkDbName value='Bak_db.mdb'>&nbsp;如备份目录有该文件，将覆盖，如没有，将自动创建<BR>"&_
"<input type='submit' value='开始备份' class='button'></form>"
ShowTable Caption,Content
If request("Go")="Start" then
	Dim fso,DbPath,BkFolder,BkDbName
	On error resume next
		DbPath=BBS.Fun.GetForm("DbPath")
		DbPath=server.mappath(DbPath)
		BkFolder=BBS.Fun.GetForm("BkFolder")
		BkDbName=BBS.Fun.GetForm("BkDbName")
		If Not IsAccess(Dbpath) Then 
			BBS.NetLog"操作后台_备份数据库失败!"
			GoBack"","备份的文件不是合法的数据库。"
			Exit Sub
		End If
		
		Set Fso=server.CreateObject("Scr"&"ipting"&".Fil"&"eSy"&"stemOb"&"ject")
		if fso.fileexists(DbPath) then
			If CheckDir(BkFolder) = True Then
			fso.copyfile DbPath,BkFolder& "\"& BkDbName
			else
			MakeNewsDir BkFolder
			fso.copyfile DbPath,BkFolder& "\"& BkDbName
			end if
			Caption="备份成功":Content="备份数据库成功!您备份的数据库路径为 " &BkFolder& "\"& BkDbName
			BBS.NetLog"操作后台_"&Content
		Else
			Caption="错误信息":Content="找不到您所需要备份的文件。"
			BBS.NetLog"操作后台_备份数据库失败!"
		End if
	ShowTable Caption,Content
End if
End sub
'---检查某一目录是否存在-----
Function CheckDir(FolderPath)
Dim Fso1
	Folderpath=Server.MapPath(".")&"\"&folderpath
    Set fso1 = CreateObject("Scr"&"ipting"&".Fil"&"eSy"&"stemOb"&"ject")
    If fso1.FolderExists(FolderPath) then
       '存在
       CheckDir = True
    Else
       '不存在
       CheckDir = False
    End if
    Set fso1 = nothing
End Function
'---根据指定名称生成目录-----
Function MakeNewsDir(foldername)
Dim fso1
	dim f
    Set fso1 = CreateObject("Scr"&"ipting"&".Fil"&"eSy"&"stemOb"&"ject")
        Set f = fso1.CreateFolder(foldername)
        MakeNewsDir = True
    Set fso1 = nothing
End Function

Sub RestoreData()
Dim Caption,Content
Caption="恢复论坛数据"
Content="<b>注意事项：</b>恢复数据库 一般是用来恢复(数据丢失或被破坏)的当前使用数据库。<br>是用备份的数据库直接把当前使用的数据库直接覆盖，请注意！<br>下面的路径都是相对与程序空间根目录的相对路径。<hr size=1 />"&_
"<form method='post' style='margin:0' action='?action=RestoreData&Go=Start'>备份数据库(相对路径)：<input type='text' size='30' name='BackPath' value='Data_Backup\Bak_db.mdb'> 请填写用来恢复的备份文件<BR>"&_
"当前数据库(相对路径)：<input type='text' size='30' name='DbPath' value='data/db.mdb'> 填写您当前使用的数据库<br /><input onclick=""if(confirm('此操作将覆盖数据库！！！\n您确定要用备份的数据库覆盖当然使用的数据库吗！？'))form.submit()"" type='button' class='button' value='恢复数据'></form> "
ShowTable Caption,Content
If request("Go")="Start" then
 Caption="错误信息"
 Dim FSO,Dbpath,BackPath
 	DbPath=BBS.Fun.GetForm("DbPath")
	BackPath=BBS.Fun.GetForm("BackPath")
	if BackPath="" or DbPath="" then
		Content="请把全名填写完整！"	
	'ElseIF Lcase(Dbpath)<>Lcase(Db) Then
		'Content="您输入的不是当前使用数据库全名!"	
	Else
	On error resume next
		DbPath=server.mappath(DbPath)
		BackPath=server.mappath(BackPath)
		
		
		If Not IsAccess(BackPath) Then
			GoBack"",Content&" 备份的文件不是合法的数据库。"
			Exit Sub
		End If
		
		Set Fso=server.CreateObject("Scr"&"ipting"&".Fil"&"eSy"&"stemOb"&"ject")
		if fso.fileexists(DbPath) then  					
		On Error Resume Next
		fso.copyfile BackPath,DbPath
			If err.number=0 then
			Caption="恢复成功":Content="成功的恢复数据库！"
			BBS.NetLog"操作后台_用"&BackPath&Content
			Else
			Content="备份目录下并无您的备份文件！"
			Err.clear
			End If
		else
		Content= "不是当前使用的数据库全名"
	End if
 End IF
ShowTable Caption,Content
End If
End sub

Sub AllUpdateGrade
	Dim ID,orders,GradeName,EssayNum,SqlEssayNum,PIC,Spic,Flag,i,S,UpdateUser,Grouping
	Grouping=Int(request.form("Grouping"))
	If Grouping=0 Then'发帖组
		For i=1 to request.form("ID").count
			ID = Replace(request.form("ID")(i),"'","")
			GradeName = Replace(request.form("GradeName")(i),"'","")
			EssayNum = Replace(request.form("EssayNum")(i),"'","")
			If GradeName="" Then
				GoBack "","等级名称必需填写!"
				Exit Sub
			End IF		
			IF Not isnumeric(ID) or Not isnumeric(EssayNum)  Then
				GoBack "","帖数需用数字填写!"
				Exit Sub
			End IF
			If EssayNum=0 Then S="OK"
		Next
		If S<>"OK" Then
			GoBack "","更新失败，你必需设置一个等级组的帖数为<font color=red>0</font>!"
			Exit Sub
		End If
	End If
	S=""
	For i=1 to request.form("ID").count
		ID = Replace(request.form("ID")(i),"'","")
		GradeName = Replace(request.form("GradeName")(i),"'","")
		PIC = Replace(request.form("PIC")(i),"'","")
		Spic = Replace(request.form("Spic")(i),"'","")
		SqlEssayNum=""
		If Grouping=0 Then
			EssayNum = Replace(request.form("EssayNum")(i),"'","")
			Set Rs=BBS.Execute("Select EssayNum FROM [Grade] where ID="&ID&" And Grouping=0")
			If Not Rs.eof Then
				If Int(EssayNum)<>Int(Rs(0)) Then
					SqlEssayNum = "EssayNum="&EssayNum&","
					S=S&ID&","
				End If
			End If
			Rs.Close
		End If
		BBS.execute("Update [Grade] Set GradeName='"&GradeName&"',"&SqlEssayNum&"PIC='"&PIC&"',Spic='"&Spic&"' where ID="&ID&" And Grouping="&Grouping)
	Next
	If Grouping=0 Then
		i=0
		If S<>"" Then
			If S<>"" THen S=left(S,len(S)-1)
			Set Rs=BBS.execute("SELECT ID,EssayNum,IsVIP from [USER] where GradeID in("&S&") And GradeFlag=0")
			Do while not Rs.eof
			I=I+1
			BBS.UpdateGrade Rs(0),Rs(1),0
			Rs.movenext
			Loop
			Rs.Close
		End IF
		S="发帖升级等级组(同时更新了"&I&"位会员)"
	ElseIF Grouping=1 Then
		S="特殊定制等级组"
	ElseIF Grouping=2 Then
		S="系统固定等级组"
	End If
	BBS.Cache.Clean("GradeInfo")
	 BBS.NetLog"操作后台_更新"&S&"成功!"
	 Suc"","更新"&S&"成功!","admin_action.asp?action=Grade"
End Sub

Sub DelGrade
	Dim ID,S,I
	ID=request("ID")
	S=BBS.execute("Select GradeName from [Grade] where ID="&ID&"")(0)
	BBS.execute("Delete * from [Grade] where ID="&ID&"")
	Set Rs=BBS.execute("Select ID,EssayNum From [User] where GradeID="&ID)
	Do while Not Rs.eof
	I=I+1
	BBS.UpdateGrade Rs(0),Rs(0),0
	Rs.movenext
	Loop
	Rs.close
	BBS.Cache.Clean("GradeInfo")
	S="删除等级组 "&S&" 成功!同时修正了"&I&"位会员"
	BBS.NetLog"操作后台_"&S
	Suc"","删除等级组 "&S,"admin_action.asp?action=Grade"
End Sub

Sub SaveGrade
Dim S,i,ID,Strings,GradeName,EssayNum,Pic,Spic,Grouping,Flag
	GradeName = BBS.Fun.GetStr("GradeName")
	EssayNum = BBS.Fun.GetStr("EssayNum")
	Pic = BBS.Fun.GetStr("Pic")
	Spic=BBS.Fun.GetStr("Spic")
	ID=Request.form("ID")
	Grouping=Request.form("Grouping")
	If GradeName="" Then GoBack "","等级名称必需填写!":Exit Sub
	Strings=BBS.Fun.GetStr("S0")&"|"
	If len(S)>8 Then GoBack"用户的颜色填写不正确","":Exit Sub
	For i=1 to 37
		IF Request.form("S"&i)="" Then GoBack"","":Exit Sub
		If Not BBS.Fun.isInteger(Request.form("S"&i)) then
			GoBack "","一些参数必须填为正整数，否则论坛不能正常运行。"
			Exit Sub
		End if
		Strings=Strings&Request("S"&i)&"|"
	Next
	Strings=Strings&"0|0|0"
	If Grouping=2 And ID<>"" Then
		BBS.execute("Update [Grade] Set GradeName='"&GradeName&"',PIC='"&PIC&"',Spic='"&Spic&"',Strings='"&Strings&"' where Grouping=2 and ID="&ID)
		S="编辑系统固定等级组成功!"
		BBS.NetLog"操作后台_"&S
		Suc"",S,"admin_action.asp?action=Grade"
	Else	
		If Grouping=0 Then
			IF Not BBS.Fun.isInteger(EssayNum)  Then GoBack "","帖数需用数字填写!":Exit Sub
			Flag=0
			S="发帖升级等级组"
		ElseIf Grouping=1 Then
			EssayNum=0
			Flag=1
			S="特殊定制等级组"
		End If
		
		If ID<>"" Then
			BBS.execute("Update [Grade] Set Grouping="&Grouping&",GradeName='"&GradeName&"',EssayNum="&EssayNum&",PIC='"&PIC&"',Spic='"&Spic&"',Strings='"&Strings&"',Flag="&Flag&" where ID="&ID)
			S="编辑"&S&"成功!"
		Else
			BBS.execute("insert into [Grade] (Grouping,GradeName,EssayNum,PIC,Spic,Flag,Strings)values("&Grouping&",'"&GradeName&"',"&EssayNum&",'"&PIC&"','"&Spic&"',"&Flag&",'"&Strings&"')")
			S="添加"&S&"成功!"
		End If
		BBS.Cache.Clean("GradeInfo")
		BBS.NetLog"操作后台_"&S
		Suc"",S&"<li>注意：在线用户要在下次重新登陆才会生效</li>","admin_action.asp?action=Grade"
	End If
End Sub

Sub AllSms
	Dim SmsContent,UserType,Sql,Mrs,I
	SmsContent=BBS.Fun.GetStr("content")
	UserType=BBS.Fun.GetStr("caption")
	If SmsContent="" Then GoBack"","":Exit Sub
	If Len(SmsContent) >3000 Then GoBack"","字符过多":Exit Sub
If UserType="1" Then
	Dim Temp,OnlineCache,Eachonline
	OnlineCache=BBS.Cache.Value("OnlineCache")
	EachOnline=Split(OnlineCache,",")
	For I=0 to uBound(EachOnline)
	Temp=Split(EachOnline(I),"|")
	BBS.Execute("insert into [sms](name,MyName,Content,MyFlag) values('论坛小信使','"&Temp(1)&"','"&SmsContent&"',1)")
	BBS.Execute("update [user] set NewSmsNum=NewSmsNum+1,SmsSize=SmsSize+1 where Name='"&Temp(1)&"'")
	BBS.UpdageOnline Temp(1),1
	Next
Else
	Select case UserType
	case"8"
	sql="select name,max(boardid) as boardid from [admin] where boardid=-1 group by name"
	Case"7"
	sql="select name,max(boardid) as boardid from [admin] where boardid>0 group by name"
	case"9"
	sql="select name,max(boardid) as boardid from [admin] where boardid=0 group by name"
	case"10"
	sql="select name,max(boardid) as boardid from [admin] group by name"
	case"4"
	sql="select name from [user] where isdel=0 and IsVip=1"
	case"0"
	sql="select name from [user] where isdel=0"
	end select
	Set Rs=BBS.Execute(Sql)
	If Not Rs.Eof Then
	MRs=Rs.GetRows(-1)
	rs.close
	For I=0 to Ubound(MRs,2)
	BBS.Execute("insert into [sms](name,MyName,Content,MyFlag) values('论坛小信使','"&MRs(0,i)&"','"&SmsContent&"',1)")
	BBS.Execute("update [user] set NewSmsNum=NewSmsNum+1,SmsSize=SmsSize+1 where Name='"&MRs(0,i)&"'")
	BBS.UpdageOnline MRs(0,i),1
	Next
	End If
End IF
Suc"","成功的群发了信件!","admin_SetHtmlEdit.asp?action=AllSms"
BBS.NetLog"操作后台_群发信件"
End Sub


Sub Bank
	Dim Coin,UserType,Sql,Mrs,I,S,Flag
	UserType=BBS.Fun.GetStr("user")
	Coin=BBS.Fun.GetStr("Coin")
	Flag=BBS.Fun.GetStr("Flag")
	If UserType="" Then GoBack"","请选择用户":Exit Sub
	If  Coin="" or Coin="0" Then GoBack"","":Exit Sub
	If Not isnumeric(Coin) Then GoBack"","请用数字填写！":Exit Sub
If UserType="1" Then
	Dim Temp,OnlineCache,Eachonline
	OnlineCache=BBS.Cache.Value("OnlineCache")
	EachOnline=Split(OnlineCache,",")
	For I=0 to uBound(EachOnline)-1
	Temp=Split(EachOnline(I),"|")
	If Flag="1" Then
		BBS.Execute("update [user] set Coin=Coin+"&Coin&" where Name='"&Temp(1)&"'")
	Else
		BBS.Execute("update [user] set Coin=Coin-"&Coin&" where Name='"&Temp(1)&"'")
	End If
	BBS.UpdageOnline Temp(1),3
	Next
Else
	Select case UserType
	case"8"
	sql="select name,max(boardid) as boardid from [admin] where boardid=-1 group by name"
	Case"7"
	sql="select name,max(boardid) as boardid from [admin] where boardid>0 group by name"
	case"9"
	sql="select name,max(boardid) as boardid from [admin] where boardid=0 group by name"
	case"10"
	sql="select name,max(boardid) as boardid from [admin] group by name"
	case"4"
	sql="select name from [user] where isdel=0 and IsVip=1"
	case"0"
	sql="select name from [user] where isdel=0"
	end select
	Set Rs=BBS.Execute(Sql)
	If Not Rs.Eof Then
	MRs=Rs.GetRows(-1)
	rs.close
	For I=0 to Ubound(MRs,2)
	If Flag="1" Then
		BBS.Execute("update [user] set Coin=Coin+"&Coin&" where Name='"&MRs(0,i)&"'")
	Else
		BBS.Execute("update [user] set Coin=Coin-"&Coin&" where Name='"&MRs(0,i)&"'")
	End If	
	BBS.UpdageOnline MRs(0,i),3
	Next
	End If
End IF
If Flag=1 Then S="送钱" Else S="扣钱"
Suc"","成功的"&S&Coin&"!","admin_action.asp?action=bank"
BBS.NetLog"操作后台_银行"&Coin
End Sub

Sub AllBoardadmin
	Dim BoardadminName,Flag,boardid,Temp,S,i,GradeFlag
	BoardadminName=BBS.Fun.GetStr("Name")
	Flag=BBS.Fun.GetStr("Flag")
	If BoardadminName="" Then GoBack"","":Exit Sub
	Set Rs=BBS.execute("Select ID,Name,password,GradeFlag,EssayNum,IsVIP From[user] where name='"&BoardadminName&"'")
	If Rs.eof Then
		GoBack"","不能操作，该用户名称还没有注册。":Exit Sub
	End If
	If Flag="Add" Then
		If not BBS.Execute("select Name From[admin] where name='"&BoardadminName&"' and boardid<1").eof Then
			GoBack"","该用户已经是超级版主或站长了。":Exit Sub
		End If
			BBS.execute("Insert into[admin](name,[password],boardid)values('"&Rs(1)&"','"&Rs(2)&"',-1)")			
			BBS.UpdateGrade Rs(0),0,8
			S="成功的添加了超级版主 "&BoardadminName&" !"
	Else
		If BBS.Execute("select Name From[admin] where name='"&BoardadminName&"' and boardid=-1").eof Then GoBack"","该用户不是超级版主":Exit Sub
		BBS.Execute("Delete From[admin] where boardid=-1 And Name='"&BoardadminName&"'")
		If Not BBS.Execute("select Name From[admin]").eof Then
			GradeFlag=7
		Else
			GradeFlag=0
			If Rs(5)=1 Then GradeFlag=4
		End If
		BBS.UpdateGrade Rs(0),Rs(3),GradeFlag
		S="成功撤消了超级版主 "&BoardadminName&" 的职位！"
	End if
	'在线刷新
	BBS.UpdageOnline BoardadminName,3
	Rs.Close
	BBS.NetLog "操作后台_"&S
	BBS.Cache.clean("BoardInfo")
	Suc"",S,"admin_action.asp?action=Boardadmin"
End Sub


Sub Boardadmin
	Dim BoardadminName,Flag,boardid,Temp,S,i
	BoardadminName=BBS.Fun.GetStr("Name")
	Flag=BBS.Fun.GetStr("Flag")
	boardid=BBS.Fun.GetStr("boardid")	
	If BoardadminName="" Then
		GoBack"",""
		Exit Sub
	ElseIf boardid="" Then
		GoBack"","请先选择管理的论坛版块"
		Exit Sub
	End If
	Set Rs=BBS.execute("Select ID,Name,password,GradeFlag,EssayNum,IsVIP From[user] where name='"&BoardadminName&"'")
	If Rs.eof Then
	GoBack"","不能添加版主，该用户名称还没有注册。":Exit Sub
	End If
	If Flag="Add" Then
		If Not BBS.Execute("select Name From[admin] where boardid="&boardid&" and Name='"&BoardadminName&"'").eof Then
			GoBack"","该用户已经是本版的版主了。":Exit Sub
		Else
			Temp=BBS.Execute("Select Boardadmin From[Board] where boardid="&boardid&"")(0)
			If Temp="" or isnull(Temp) Then
				Temp=BoardadminName
			Else
				Temp=Temp&"|"&BoardadminName
			End If
			BBS.execute("Insert into[admin](name,[password],boardid)values('"&Rs(1)&"','"&Rs(2)&"',"&boardid&")")
			BBS.execute("update [Board] Set Boardadmin='"&Temp&"' where boardid="&boardid)
			
		If Rs(3)<7 Then
		BBS.UpdateGrade Rs(0),0,7
		End If
			S="成功的添加了版主 "&BoardadminName&" !"
		End If
	Else
		Temp=BBS.Execute("Select Boardadmin From[Board] where boardid="&boardid&"")(0)
		Temp=split(Temp,"|")
		For i=0 To uBound(temp)
			IF lcase(BoardadminName)<>lcase(Temp(i)) Then S=S&Temp(i)&"|"
		Next
		If S<>"" Then S=left(S,len(S)-1)
		BBS.Execute("Delete From[admin] where boardid="&boardid&" And Name='"&BoardadminName&"'")
		BBS.Execute("Update [Board] Set Boardadmin='"&S&"' where boardid="&boardid&"")
		If Rs(3)=7 Then
			If Rs(5)=1 Then'如果是VIP
				BBS.UpdateGrade Rs(0),0,4
			Else
				BBS.UpdateGrade Rs(0),Rs(3),0
			End If
		End If
		S="成功撤消了版主 "&BoardadminName&" 的职位！"
	End if
	Rs.Close
	'在线刷新
	BBS.UpdageOnline BoardadminName,3
	BBS.NetLog "操作后台_"&S
	BBS.Cache.clean("BoardInfo")
	Suc"",S,"admin_action.asp?action=Boardadmin"
End Sub

Sub DelFaction
	Dim Name
	Name=Request.QueryString("Name")
	BBS.Execute("Delete * From[Faction] where Name='"&Name&"'")
	BBS.Execute("update [User] Set Faction='' where Faction='"&Name&"'")
	BBS.NetLog "操作后台_删除帮派 "&Name
	Suc"","成功删除了帮派！","admin_action.asp?action=Faction"
End Sub

Sub SaveFaction
	Dim Name,FullName,Note,User,BuildDate,ID,S
	ID=BBS.Fun.GetStr("ID")
	Name=BBS.Fun.GetStr("Name")
	FullName=BBS.Fun.GetStr("FullName")
	Note=BBS.Fun.GetStr("Note")
	User=BBS.Fun.GetStr("User")
	BuildDate=BBS.Fun.GetStr("BuildDate")
	IF Name="" Or FullName="" Or Note="" or User="" Then Call Goback("",""):Exit Sub
	If Not isDate(BuildDate) Then BuildDate=BBS.NowBBSTime
	IF BBS.Execute("select name From[User] where Name='"&User&"'").eof Then
		GoBack"","掌门人必须是注册会员！":Exit Sub
	End If
	If ID<>"" Then
		Set Rs=BBS.Execute("Select Name From[Faction] where ID="&ID)
		If Rs.Eof Then Goback"","记录已被删除了！"
		If Name<>Rs(0) Then
		BBS.Execute("update [User] Set Faction='"&Name&"' where Faction='"&Rs(0)&"'")
		End If
		Rs.Close
		BBS.Execute("update [Faction] Set [Name]='"&Name&"',FullName='"&FullName&"',[Note]='"&Note&"',[User]='"&User&"',BuildDate='"&BuildDate&"' where ID="&ID)
		S="成功修改了帮派："&Name
	Else
		BBS.execute("Insert into[Faction](Name,FullName,[Note],BuildDate,[User])Values('"&Name&"','"&FullName&"','"&Note&"','"&BuildDate&"','"&User&"')")
		BBS.Execute("update [User] Set Faction='"&Name&"' where Name='"&User&"'")
		S="成功添加了帮派："&Name
	End If
		Suc"",S,"admin_action.asp?action=Faction"
	BBS.NetLog "操作后台_"&S
End Sub

Sub DelEssay
	Dim UserName,DateNum,boardid,AllTable,I,SqlWhere,S
	DateNum=BBS.Fun.GetStr("DateNum")
	boardid=BBS.Fun.GetStr("boardid")
	UserName=BBS.Fun.GetStr("Name")
	AllTable=Split(BBS.BBStable(0),",")
	Select Case Request("Go")
	Case"Date"
		If not isnumeric(DateNum) Then GoBack"","天数必需用数字填写！":Exit Sub
		If boardid=0 Then
			SQlwhere=""
			S="成功删除"&DateNum&"天前发表的主题帖（包括其回复帖）！"
		Else
			S="成功删除在 "&BBS.Execute("Select BoardName From[Board]where boardid="&boardid&"")(0)&" 上"&DateNum&"天前发表的主题帖（包括其回复帖）！"
			Sqlwhere=" And boardid="&boardid
		End If
		Set Rs=BBS.Execute("Select TopicID,SqlTableID,boardid From [Topic] where DATEDIFF('d',[AddTime],'"&BBS.NowBbsTime&"')>"&DateNum&Sqlwhere&"")
		Do while not Rs.eof
		BBS.Execute("Delete From[Bbs"&Rs(1)&"] where (TopicID="&Rs(0)&" or ReplyTopicID="&Rs(0)&") "&SqlWhere&"")
		BBS.Execute("Delete From[Appraise] where TopicID="&Rs(0))
		BBS.execute("Delete From [TopicVote] where TopicID="&Rs(0))
		BBS.execute("Delete From [TopicVoteUser] where TopicID="&Rs(0))
		Rs.MoveNext
		Loop
		Rs.close
		BBS.Execute("Delete From[Topic] where DATEDIFF('d',[AddTime],'"&BBS.NowBbsTime&"')>"&DateNum&Sqlwhere)
		Suc"",S&"<li>删除后建议对论坛做一次<a href=admin_action.asp?action=UpdateBbs>整理</a>，来重计版面的帖数","admin_action.asp?action=DelEssay"
		BBS.NetLog"操作后台_"&S	
	Case"DateNoRe"
		If not isnumeric(DateNum) Then GoBack"","天数必需用数字填写！":Exit Sub
		If boardid=0 Then
			SQlwhere=""
			S="成功删除"&DateNum&"天前没有回复的所有主题帖（包括其回复）！！"
		Else
			S="成功删除在 "&BBS.Execute("Select BoardName From[Board]where boardid="&boardid&"")(0)&" 上"&DateNum&"天前没有回复的所有主题帖（包括其回复帖）！"
			Sqlwhere=" And boardid="&boardid
		End If
		Set Rs=BBS.Execute("Select TopicID,SqlTableID,boardid From [Topic] where DATEDIFF('d',LastTime,'"&BBS.NowBbsTime&"')>"&DateNum&Sqlwhere&"")
		Do while not Rs.eof
			BBS.Execute("Delete From[Bbs"&Rs(1)&"] where (TopicID="&Rs(0)&" or ReplyTopicID="&Rs(0)&") "&SqlWhere&"")
			BBS.Execute("Delete From [Appraise] where TopicID="&Rs(0))
			BBS.execute("Delete From [TopicVote] where TopicID="&Rs(0))
			BBS.execute("Delete From [TopicVoteUser] where TopicID="&Rs(0))
		Rs.MoveNext
		Loop
		Rs.close
		BBS.Execute("Delete From [Topic] where DATEDIFF('d',LastTime,'"&BBS.NowBbsTime&"')>"&DateNum&Sqlwhere)
		Suc"",S&"<li>删除后建议对论坛做一次<a href=admin_action.asp?action=UpdateBbs>整理</a>，来重计版面的帖数","admin_action.asp?action=DelEssay"
		BBS.NetLog"操作后台_"&S	
	Case"User"
		If UserName="" Then GoBack"","":Exit Sub
		IF BBS.Execute("select name From[User] where Name='"&UserName&"'").eof Then
			GoBack"","这个用户根本不存在！":Exit Sub
		End If
		If boardid=0 Then
			SQlwhere=""
			S="成功删除用户 "&UserName&" 的所有帖子！！"
		Else
			S="成功删除用户 "&UserName&"在 "&BBS.Execute("Select BoardName From[Board]where boardid="&boardid&"")(0)&" 的帖子！"
			Sqlwhere=" And boardid="&boardid
		End If
			Set Rs=BBS.Execute("select TopicID,SqltableID From[Topic] where Name='"&UserName&"'"&SqlWhere&"")
			do while not Rs.eof
				BBS.Execute("Delete From[Bbs"&Rs(1)&"] where (TopicID="&Rs(0)&" or ReplyTopicID="&Rs(1)&") "&SqlWhere&"")
				BBS.Execute("Delete From [Appraise] where TopicID="&Rs(0))
				BBS.execute("Delete From [TopicVote] where TopicID="&Rs(0))
				BBS.execute("Delete From [TopicVoteUser] where TopicID="&Rs(0))
			Rs.movenext
			Loop
			BBS.Execute("Delete From[Topic] where Name='"&UserName&"'"&SqlWhere&"")		
			For i=0 To uBound(AllTable)
			BBS.Execute("Delete From[Bbs"&AllTable(i)&"] where Name='"&UserName&"'"&SqlWhere&"")
			Next
			Suc"",S&"<li>删除后建议对论坛做一次<a href=admin_action.asp?action=UpdateBbs>整理</a>，来重计版面的帖数","admin_action.asp?action=DelEssay"
		BBS.NetLog"操作后台_"&S
	Case Else
	GoBack"","提交的路径不正确"
	End Select
End Sub

Sub DelSms
	Dim UserName,DateNum,boardid,S
	DateNum=BBS.Fun.GetStr("DateNum")
	Select Case Request("Go")
	Case"Date"
		If not isnumeric(DateNum) Then GoBack"","天数必需用数字填写！":Exit Sub
		BBS.Execute("Delete From[Sms] where DATEDIFF('d',[AddTime],'"&BBS.NowBbsTime&"')>"&DateNum&"")
		S="已经成功删除在"&DateNum&"天前的所有留言信件！"
	Case"User"
		UserName=BBS.Fun.GetStr("Name")
		IF UserName="" Then GoBack"","":Exit Sub
		IF BBS.Execute("select name From[User] where Name='"&UserName&"'").eof Then GoBack"","这个用户根本不存在！":Exit Sub
		BBS.Execute("Delete From[Sms] where MyName='"&UserName&"'")
		S="成功删除了用户 "&UserName&" 的所有留言信件！"
	Case"Auto"
		If not isnumeric(DateNum) Then GoBack"","天数必需用数字填写！":Exit Sub
		BBS.Execute("Delete From[Sms] where DATEDIFF('d',[AddTime],'"&BBS.NowBbsTime&"')>"&DateNum&" And Name In ('自动送信系统','论坛小信使')")
		S="成功删除在"&DateNum&"天前的所有论坛自动送信的留言信件！"
	End Select
	BBS.NetLog "操作后台_"&S
	Suc"",S,"admin_action.asp?action=DelSms"
End Sub

Sub MoveEssay
	Dim boardid1,boardid2,DateNum,UserName,AllTable,I,S
	boardid1=BBS.Fun.GetStr("boardid1")
	boardid2=BBS.Fun.GetStr("boardid2")
	IF boardid1=boardid2 Then GoBack"","您还没有选择目标论坛！":Exit Sub
	AllTable=Split(BBS.BBStable(0),",")
	DateNum=BBS.Fun.GetStr("DateNum")
	UserName=BBS.Fun.GetStr("Name")
Select Case Request("Go")
Case"Date"
	If not isnumeric(DateNum) Then GoBack"","天数必需用数字填写！":Exit Sub
	Set Rs=BBS.Execute("Select TopicID,SqlTableID from[Topic] Where DATEDIFF('d',[AddTime],'"&BBS.NowBbsTime&"')>"&DateNum&"  and boardid="&boardid1) 
	Do while not Rs.eof
	BBS.Execute("update [Bbs"&Rs(1)&"] Set boardid="&boardid2&" where boardid="&boardid1&" and (TopicID="&Rs(0)&" or ReplyTopiciD="&Rs(0)&")")
	Rs.movenext
	Loop
	Rs.Close
	BBS.Execute("update [Topic] Set boardid="&boardid2&" where boardid="&boardid1&" And DATEDIFF('d',[AddTime],'"&BBS.NowBbsTime&"')>"&DateNum&"")
	S="成功的把"&DateNum&"天前的帖子从 "
Case"User"
	If UserName="" Then GoBack"","":Exit Sub
	IF BBS.Execute("select name From[User] where Name='"&UserName&"'").eof Then
		GoBack"","这个用户根本不存在！":Exit Sub
	End IF
	Set Rs=BBS.Execute("Select TopicID,SqlTableID from[Topic] Where Name='"&UserName&"' And boardid="&boardid1) 
	Do while not Rs.eof
	BBS.Execute("update [Bbs"&Rs(1)&"] Set boardid="&boardid2&" where boardid="&boardid1&" and (TopicID="&Rs(0)&" or ReplyTopiciD="&Rs(0)&")")
	Rs.movenext
	Loop
	Rs.Close
	BBS.Execute("update [Topic] Set boardid="&boardid2&"  Where boardid="&boardid1&" and Name='"&UserName&"'")
	S="成功的把"&UserName&"的帖子从 "
End Select
	S=S&BBS.Execute("select BoardName From[Board] where boardid="&boardid1&"")(0)&" 移动到 "&BBS.Execute("select BoardName From[Board] where boardid="&boardid2&"")(0)&"！"
	BBS.NetLog"操作后台_"&S	
	Suc"",S&"！现在进行一次 <a href='admin_Board.asp?action=BoardUpdate'>版面整理</a> ！","admin_action.asp?action=MoveEssay"
End Sub

Sub Clean
	Application.Contents.RemoveAll
	Suc "","更新缓存成功","admin_action.asp?action=Clean"
	BBS.NetLog"操作后台_更新缓存"
End Sub

Sub Topadmin
	Dim TopadminName,Flag,S,GradeFlag
	TopadminName=Replace(Request("Name"),"'","")
	Flag=Request("Flag")
	If TopadminName="" Then
		GoBack"","":Exit Sub
	End If
	Set Rs=BBS.execute("Select Name,password,ID,IsVip,EssayNum From[user] where name='"&TopadminName&"'")
	If Rs.eof Then GoBack"","该用户名称还没有注册。":Exit Sub
	If Flag="1" Then
		If Not BBS.Execute("select Name From[admin] where boardid<1 and Name='"&TopadminName&"'").eof Then
			GoBack"","该用户已经是管理员！":Exit Sub
		End If
		BBS.execute("Insert into[admin](name,[password],boardid)values('"&Rs(0)&"','"&Rs(1)&"',0)")
		BBS.UpdateGrade Rs(2),0,9
		S="成功添加了"
	Else
		BBS.Execute("delete * from [admin] where name='"&TopadminName&"' and boardid=0")
		GradeFlag=0
		If Rs(3)=1 Then Flag=4
		If Not BBS.Execute("select boardid from [admin] where name='"&TopadminName&"'").eof Then
			GradeFlag=7
		End if
		BBS.UpdateGrade Rs(2),Rs(4),GradeFlag
	 	S="成功撒销了"
	End If
	Rs.Close
	S=S&BBS.GetGradeName(0,9)&":"&TopadminName&" !"
	BBS.UpdageOnline TopadminName,3
	BBS.NetLog"操作后台_"&S
	Suc"",S,"admin_action.asp?action=Topadmin"
End Sub

Sub GapAd
	Dim content
	Dim Temp,I,objFSO,objname,TmpStr,ad_num,ad_tmp,adv_num,ii
	Set objFSO = Server.CreateObject("Scr"&"ipting"&".Fil"&"eSy"&"stemOb"&"ject")
	Set objName=objFSO.OpenTextFile(Server.MapPath("inc/ads.js"))
	tmpstr=objName.readall
	temp=split(tmpstr,chr(13)&chr(10))
	ad_num=replace(Temp(1),";if(a==0){a=1}","")
	ad_num=Int(replace(ad_num,"a=",""))
	objName.close
	Set objFSO=nothing
	Content=""
	ii=0
	for i=1 to ad_num+1
		ad_tmp=Replace(Request.form("ad_v"&i&""),"'","")
		if Trim(ad_tmp)<>"" or isnull(ad_tmp)then
		ii=ii+1
		Content=Content&"b["&ii&"].under='<img src=images/icon/ad_icon.gif align=absmiddle> "&ad_tmp&"'"&vbcrlf
		end if
	next
	Set objFSO = Server.CreateObject("Scr"&"ipting"&".Fil"&"eSy"&"stemOb"&"ject")
	Set objname=objFSO.CreateTextFile(Server.MapPath("Inc/Ads.js"),True)
	objname.Write"<!--"&vbcrlf&"a="&ii&";if(a==0){a=1}"&vbcrlf&"var slump=Math.random();"&vbcrlf&"var talet=Math.round(slump * (a-1))+1;"&vbcrlf&"function create(){ "&vbcrlf&"this.under=''"&vbcrlf&"}"&vbcrlf&"b=new Array()"&vbcrlf&"for(var i=1; i<=a; i++){b[i]=new create()}"&vbcrlf&Content&"var visa="""";"&vbcrlf&"document.write(b[talet].under+'</center>');"&vbcrlf&"//-->"
	objname.close
	set objfso=nothing
	BBS.NetLog"操作后台_修改帖间广告"
	Response.redirect"admin_action.asp?action=GapAd"
End Sub

'检验是否是数据库
Function IsAccess(AccessPath)
On Error Resume Next
IsAccess=False
Dim TempConn
	Set TempConn=Server.CreateObject("Adodb.connection")
		TempConn.Open "Provider=Microsoft.jet.oledb.4.0;data source="&AccessPath
		If Err.Number<>0 Then 
			IsAccess=False
		Else
			IsAccess=True
		End If
TempConn.Close
Set TempConn=Nothing
End Function
%>