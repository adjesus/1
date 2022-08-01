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
	If not isnumeric(BBS.Fun.GetStr("info"&i)) Then GoBack"","������������������д��":Exit Sub
	End If
	Else
		IF BBS.Fun.GetStr("info"&i)="" Then GoBack"",i:Exit Sub
		If Not BBS.Fun.isInteger(BBS.Fun.GetStr("info"&i)) then
			GoBack"","һЩ����������Ϊ��������������̳�����������С�"&i
			Exit Sub
		End if
	End IF
		S=S&Replace(Request.form("info"&i),",","&#44")&","
	Next
	S=S&",0,0,0,0,0"
	BBS.execute("update [Config] set Info='"&S&"'")
	S="�ɹ�������̳��Ϣ����"
	BBS.NetLog"������̨_"&S
	Suc"�޸ĳɹ�",S,"admin_action.asp?action=BbsInfo"
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
		GoBack"","һЩ����������Ϊ��������������̳�����������С�"
	End if
	.Execute("update [Config] set Hits="&hits&",UserNum="&UserNum&",AllEssayNum="&AllEssayNum&",TopicNum="&TopicNum&",MaxEssayNum="&MaxEssayNum&",MaxOnlineNum="&MaxOnlineNum&",MaxOnlineTime='"&MaxOnlineTime&"'")
	S="��̳ϵͳ����ͳ���޸ĳɹ�"
	Suc"�޸ĳɹ�",S,"admin_action.asp?action=ConfigData"
	.NetLog"������̨_"&S
	.Cache.Clean("InfoUpdate")
	.Cache.Clean("Hits")
	End With
End Sub

Sub DelLog()
	If Request.Form("Del")="�����־" Then
		BBS.Execute("Delete From [Log] where DATEDIFF('d', LogTime,'"&BBS.NowBBSTime&"')>2")
	Else
	Dim ID
		ID=Request.form("ID")
		If ID="" Then Goback "","����ѡ��":Exit Sub
		BBS.Execute("Delete From [Log] where ID in("&ID&") And DATEDIFF('d', LogTime,'"&BBS.NowBBSTime&"')>2")
	End If
	BBS.NetLog"������̨_��־ϵͳ-"&Request.Form("Del")
	Suc "","ɾ����־�ɹ�!ϵͳ���Զ������������־��¼��","admin_actionList.asp?action=Log"
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
		Goback"","���û���û��ע�ᣡ"
		Exit Sub
		Else
		ID=Rs(0)
		End If
		Rs.Close
	End If
	If ID="" Then Goback "","����ѡ���û�":Exit Sub
	If Point="" Then Goback"","�㻹û��ѡ����ν��в���":Exit Sub
	Set Rs=BBS.Execute("Select Name,IsVIP,IsDel,ID,GradeID,GradeFlag,EssayNum From [User] where ID in("&ID&")")
	Select case int(Point)
	Case 1
		S="���û���"
		Do while not Rs.eof
		S=S&"��"&Rs(0)&"��"
		Rs.movenext
		Loop
		S=S&"��ɾ�����!"
		BBS.Execute("update [User] Set IsDel=1 where ID in("&ID&")")
	Case 2
		S="��ȫɾ���û���"
		Do while not Rs.eof
			AllTable=Split(BBS.BBStable(0),",")
			For i=0 To uBound(AllTable)
				BBS.Execute("Delete From[Bbs"&AllTable(i)&"] where Name='"&Rs(0)&"'")
				BBS.Execute("Delete From[Bbs"&AllTable(i)&"] where ReplyTopicID in (Select TopicID From[Topic] where Name='"&Rs(0)&"')")	
			Next
			BBS.Execute("Delete From[Topic] where  Name='"&Rs(0)&"'")
			BBS.Execute("Delete From[Sms] where  MyName='"&Rs(0)&"'")
			BBS.Execute("Delete From[admin] where Name='"&Rs(0)&"'")
			S=S&"��"&Rs(0)&"��"
		Rs.movenext
		Loop
		BBS.Execute("Delete * From [User] where ID in("&ID&")")
	Case 3
		S="����ɾ���û���"
		Do while not Rs.eof
			AllTable=Split(BBS.BBStable(0),",")
			For i=0 To uBound(AllTable)
				BBS.Execute("Delete From[Bbs"&AllTable(i)&"] where Name='"&Rs(0)&"'")
				BBS.Execute("Delete From[Bbs"&AllTable(i)&"] where ReplyTopicID in (Select TopicID From[Topic] where Name='"&Rs(0)&"')")	
			Next
			BBS.Execute("Delete From[Topic] where Name='"&Rs(0)&"'")
			S=S&"��"&Rs(0)&"��"
		Rs.movenext
		Loop
		S=S&"���������ӳɹ�!"			
	Case 4
		S="�����û���"
		Do while not Rs.eof
		S=S&"��"&Rs(0)&"��"
		Rs.movenext
		Loop
		S=S&"���������ӳɹ�!"
		BBS.Execute("update [User] Set IsShow=1 where ID in("&ID&")")
	Case 5
		S="�����û���"
		Do while not Rs.eof
		S=S&"��"&Rs(0)&"��"
		Rs.movenext
		Loop
		S=S&"�ĸ���ǩ���ɹ�!"
		BBS.Execute("update [User] Set IsSign=1 where ID in("&ID&")")
	Case 6
		S="�����û���"
		Do while not Rs.eof
		S=S&"��"&Rs(0)&"��"
		If Rs(5)=0 Then'�������ͨ�ȼ�����
			BBS.UpdateGrade Rs(3),0,4
		End If
		Rs.movenext
		Loop
		S=S&"ΪVIP��Ա!"
		BBS.Execute("update [User] Set IsVip=1 where ID in("&ID&")")	
	Case 7
		Dim EssayNum,GoodNum,Grade,Rs1
		S="�����޸��û���"
		AllTable=Split(BBS.BBStable(0),",")
		Do while not rs.eof
		S=S&"��"&Rs(0)&"��"
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
		S=S&"�������ݳɹ�!"
	Case 8
		Dim GradeID
		GradeID=Request.form("GradeID")
		If GradeID="" Then Goback"",""
		S="�����û���"
		Do while not Rs.eof
		If Rs(5)<4 Then'�԰������ϼ�����Ч
			S=S&"��"&Rs(0)&"��"
			IsBe=True
			BBS.execute("update [User] set GradeID="&GradeID&",GradeFlag=1 where Name='"&Rs(0)&"'")
			BBS.UpdageOnline Rs(0),3
		End If
		Rs.movenext
		Loop
		If Not IsBe Then Goback"","ѡ�����û��Ѿ��ǰ������ϵĵȼ���":Exit Sub
		S=S&"Ϊ�ر�ȼ��� "&BBS.GetGradeName(GradeID,0)&" �ɹ�!"
	Case 9
		S="���û���"
		Do while not Rs.eof
		If Rs(5)=1 Then
			IsBe=True
			S=S&"��"&Rs(0)&"��"
			If Rs(1)=1 Then
				BBS.UpdateGrade Rs(3),0,4
			Else
				BBS.UpdateGrade Rs(3),Rs(6),0
			End If
			BBS.UpdageOnline Rs(0),3
		End If
		Rs.movenext
		Loop
		If Not IsBe Then Goback"","ѡ�����û��������ر�ȼ���":Exit Sub
		S=S&"�������������ȼ���ɹ�!"
	Case 10	
		S="ͨ��ע���û���"
		Do while not Rs.eof
		S=S&"��"&Rs(0)&"��"
		Rs.movenext
		Loop
		S=S&"�����!"
		BBS.Execute("update [User] Set IsDel=0 where ID in("&ID&")")
	Case 12	
		S="��ɾ���û���"
		Do while not Rs.eof
		S=S&"��"&Rs(0)&"��"
		Rs.movenext
		Loop
		S=S&"�ָ��ɹ�!"
		BBS.Execute("update [User] Set IsDel=0 where ID in("&ID&")")
	Case 13
		S="���û���"
		Do while not Rs.eof
		S=S&"��"&Rs(0)&"��"
		Rs.movenext
		Loop
		S=S&"�������������ӻָ��ɹ�!"
		BBS.Execute("update [User] Set IsShow=0 where ID in("&ID&")")
	Case 14
		S="�ָ���ʾ�û���"
		Do while not Rs.eof
		S=S&"��"&Rs(0)&"��"
		Rs.movenext
		Loop
		S=S&"�ĸ���ǩ���ɹ�!"
		BBS.Execute("update [User] Set IsSign=0 where ID in("&ID&")")
	Case 11
		S="ȡ��VIP��Ա�û���"
		Do while not Rs.eof
		S=S&"��"&Rs(0)&"��"
		If Rs(5)=4 Then'�����VIP�ȼ����
		BBS.UpdateGrade Rs(3),EssayNum,0
		End If
		Rs.movenext
		Loop
		S=S&"!"
		BBS.Execute("update [User] Set IsVip=0 where ID in("&ID&")")	
	End Select
	Rs.Close
	BBS.NetLog"������̨_"&S
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
		If Subs>0 And Int(ParenID)>0 Then GoBack"","�ò˵����������˵���Ŀ��������Ϊ�����˵��":Exit Sub
		If Flag<>"" Then
			If Flag="8" Then
				BBS.Execute("Update [Menu] Set name='"&MenuName&"',Show="&Show&" where ID="&ID)
			Else
				BBS.Execute("Update [Menu] Set name='"&MenuName&"',Show="&Show&",Target="&Target&",ParenID="&ParenID&" where ID="&ID)
			End IF
		Else
			BBS.Execute("Update [Menu] Set name='"&MenuName&"',url='"&MenuUrl&"',Show="&Show&",Target="&Target&",ParenID="&ParenID&" where ID="&ID)
		End if
		S="�޸Ĳ˵���"&MenuName&" �ɹ�!"
		BBS.NetLog"������̨_"&S
		Suc"",S,"admin_action.asp?action=Menu"
	Else
		Dim Orders
		Orders=BBS.Execute("Select Count(*) from [Menu] where ParenID="&ParenID)(0)+1
		BBS.Execute("insert into [Menu](name,url,Target,Show,ParenID,Orders) values('"&MenuName&"','"&MenuUrl&"',"&Target&","&Show&","&ParenID&","&Orders&")")
		S="��Ӳ˵���"&MenuName&" �ɹ�!"
		BBS.NetLog"������̨_"&S
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
	S="�˵�������³ɹ�!"
	BBS.NetLog"������̨_"&S
	Suc "",S,"admin_action.asp?action=Menu"
End Sub

Sub DelMenu
	Dim ID,S
	ID=Request.QueryString("ID")
	BBS.Execute("Delete From[Menu] where Flag=0 And ID="&ID)
	S="ɾ���˵��ɹ�"
	BBS.NetLog"������̨_"&S
	Suc"",S,"admin_action.asp?action=Menu"
End Sub


'==--->>>����JS�ļ�
Sub SetJsMenu
	Dim objFSO,objName
	Dim UserMenu,TouristMenu,uM,tM
	Dim Target,Target1,S
	Dim Rs,Rs1,I,II
	Dim Board_Rs,Po,BoardMenu,BoardSelect
	'���ɶ����˵�
	Set Rs=BBS.Execute("Select ID,Name,Url,show,orders,flag,Target From [Menu] where show<3 and parenID=0 order by orders")
	Do while not Rs.eof
	 '���ָ��Ϊ8
	 If Rs(5)=8 Then
			S="<div class=menuitems><A href=cookies.asp?action=style&skinid=0>Ĭ�Ϸ��<\/A><\/div>"
			Set Rs1=BBS.Execute("Select skinid,SkinName From[Skins] where pass=1")
			do while Not Rs1.Eof
				S=S&"<div class=menuitems><A href=cookies.asp?action=style&skinid="&Rs1(0)&">"&BBS.Fun.GetJsStr(Rs1(1))&"<\/A><\/div>"
			rs1.movenext
			Loop
			Rs1.close
			S="<DIV id=M"&Rs(0)&" class=menu>"&S&"<\/DIV>��<a href='#' onmouseover=\""dropdownmenu(this, event, \'M"&Rs(0)&"\');\"" >"&BBS.Fun.GetJsStr(Rs(1))&"<\/a>"
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
				If Rs(3)<>2 Then UserMenu=UserMenu&"��<a>"&BBS.Fun.GetJsStr(Rs(1))&"<\/a>"
				If Rs(3)<>1 Then TouristMenu=TouristMenu&"��<a>"&BBS.Fun.GetJsStr(Rs(1))&"<\/a>"
			Else
				If Rs(3)<>2 Then UserMenu=UserMenu&"��<a href="&BBS.Fun.GetJsStr(Rs(2))&""&Target&">"&BBS.Fun.GetJsStr(Rs(1))&"<\/a>"
				If Rs(3)<>1 Then TouristMenu=TouristMenu&"��<a href="&BBS.Fun.GetJsStr(Rs(2))&""&Target&">"&BBS.Fun.GetJsStr(Rs(1))&"<\/a>"
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
				S="��<a onmouseover=\""dropdownmenu(this, event, \'M"&Rs(0)&"\');\"" href="&BBS.Fun.GetJsStr(Rs(2))&""&Target&">"&BBS.Fun.GetJsStr(Rs(1))&"<\/a>"
			Else
				S="��<a href='#' onmouseover=\""dropdownmenu(this, event, \'M"&Rs(0)&"\');\"">"&BBS.Fun.GetJsStr(Rs(1))&"<\/a>"
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
	UserMenu=Replace(UserMenu,"{�û���}","""+user+""")
'���ɰ��˵�
	Set Rs=BBS.Execute("Select Depth,boardid,ParentID,Boardname,BoardImg,Introduce,Boardadmin,PassUser,Child,ParentStr,RootID,Strings From[Board] order by RootID,Orders")
	If not Rs.Eof Then
		Board_Rs=Rs.GetRows(-1)
	End If
	If IsArrAy(Board_Rs) Then 
		For i=0 To Ubound(Board_Rs,2)
		Po=""
		If Board_Rs(0,i)=0 Then'��
			BoardMenu=BoardMenu&"<div class=\""menuitems\""><a href=\""board.asp?boardid="&Board_Rs(1,i)&"\""><b>"&BBS.Fun.GetJsStr(Board_Rs(3,i))&"</b></a></div>"
			BoardSelect=BoardSelect&"<option><b>"&Board_Rs(3,i)&"</b></option>"
		Else
			For II=2 to Board_Rs(0,i)
			Po=Po&"�O"
			Next
			BoardMenu=BoardMenu&"<div class=\""menuitems\""><A href=\""board.asp?boardid="&Board_Rs(1,i)&"\"">"&po&"&nbsp;&nbsp;�� "&BBS.Fun.GetJsStr(Board_Rs(3,i))&"</a></div>"
			BoardSelect=BoardSelect&"<option value=\"""&Board_Rs(1,i)&"\"">"&po&"&nbsp;&nbsp;�� "&Board_Rs(3,i)&"</option>"
		End IF
		Next
		BoardSelect="<select onchange=if(this.options[this.selectedIndex].value!=''){location='board.asp?boardid='+this.options[this.selectedIndex].value;}><option selected>��ת��̳��...</option>"&BoardSelect&"</select>"
		BoardMenu="<div id=\""Board\"" class=\""menu\"">"&BoardMenu&"</div>"
	End If
	
	On Error Resume Next
	Set objFSO = Server.CreateObject("Scr"&"ipting.Fil"&"eSystemOb"&"ject")
	Set objName=objFSO.CreateTextFile(Server.MapPath("Inc/setmenu.js"),True)
	objName.Write"function UserMenu(user){"&vbcrlf&"document.write("""&UserMenu&""");"&vbcrlf&"}"&vbcrlf&"function TouristMenu(){"&vbcrlf&"document.write("""&TouristMenu&""");"&vbcrlf&"}"&vbcrlf&"function BoardListMenu(){"&vbcrlf&"document.write("""&BoardMenu&""");"&vbcrlf&"}"&vbcrlf&"function BoardSelect(){"&vbcrlf&"document.write("""&BoardSelect&""");"&vbcrlf&"}"
	objName.Close
	Set objFso=Nothing
	If Err Then
		Goback"","����ʧ�ܣ��ռ䲻֧��FOS�ļ���д����"
		err.Clear
		Exit Sub
	End If
	Suc "","�ɹ��ĸ�����ǰ̨�Ĳ˵�:��������˵�����������˵�!","javascript:history.go(-1)"
	BBS.Netlog "������̨_����ǰ̨Js�˵��ɹ�!"
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
	Goback"","�޸�ʧ�ܣ��ռ䲻֧��FOS�ļ���д����ֱ�ӱ༭inc/agreement.html����ļ���"
	err.Clear
	Exit Sub
	End If
	S="�޸���̳ע��Э��ɹ�!"
	BBS.Netlog "������̨_"&S
	Suc "�޸ĳɹ�",S,"admin_SetHtmlEdit.asp?action=Agreement"
End Sub

Sub DelPlacard
	Dim ID,S
	ID=Request.QueryString("ID")
	BBS.execute("Delete From [Placard] where ID="&ID&"")
	BBS.Cache.clean("Placard")
	S="ɾ������ɹ�!"
	BBS.NetLog"������̨_"&S
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
		S="����༭�ɹ�"
	Else
		BBS.execute("insert into[Placard](Caption,Content,AddTime,[Name],boardid,Hits,UbbString)values('"&Caption&"','"&Content&"','"&AddTime&"','"&Name&"',"&B_ID&","&Hits&",'"&S&"')")
		S="���淢���ɹ�"
	End If
	.NetLog"������̨_"&S
	.Cache.clean("Placard")
	Suc"",S&"��","admin_actionList.asp?action=Placard"
	End With
End Sub

Sub DelLockIP
	Dim ID
	ID=Int(Request("ID"))
	BBS.Execute("Delete From [LockIP] Where ID="&ID&"")
	BBS.Cache.clean("IPData")
	BBS.NetLog"������̨_ɾ����IP��¼�ɹ�!"
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
		S="�޸ķ���IP�ɹ�!"
		Suc"",S,"admin_action.asp?action=LockIp"
	Else
		BBS.execute("Insert into [LockIp](StartIp,EndIp,Readme,lock)values("&BBS.Fun.IpEnCode(StartIp)&","&BBS.Fun.IpEnCode(EndIp)&",'"&Readme&"',1)")
		S="���η����ɹ�!"
		Suc"���η����ɹ�","��ù�ĺ��ӵ�IP�Ѿ�����!","admin_action.asp?action=LockIp"
	End If
	BBS.NetLog"������̨_"&S
	BBS.Cache.clean("IPData")
End Sub

Sub IsLockIp
	Dim ID,IsLock,S
	ID=Int(Request("ID"))
	IsLock=BBS.Execute("Select Lock From[LockIp] where Id="&ID&"")(0)
	If IsLock=1 Then
		S="���IP�ɹ�!"
		BBS.Execute("update [LockIp] set Lock=0 where Id="&ID&"") 
	Else
		S="����IP�ɹ�!"
		BBS.Execute("update [LockIp] set Lock=1 where Id="&ID&"") 
	End IF
	BBS.NetLog"������̨_"&S
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
	If S="" Then Goback"ϵͳ����","��Ч�����ݱ����ƣ�":Exit Sub
	IF Int(Aute)<>Int(BBS.BBStable(1)) Then
		S=BBS.BBStable(0)&"|"&Int(Aute)
		BBS.execute("Update [Config] Set BbStable='"&S&"' ")
	End If
	S="������̳Ĭ�����ݱ�Ϊ bbs"&Aute&" �ɹ�!"
	BBS.NetLog"������̨_"&S
	Suc"",S,"admin_action.asp?action=SqlTable"
	BBS.Cache.clean("parameter")
End Sub

Sub AddSqlTable
	Dim TableName,AllTable,I,S
	TableName=BBS.Fun.GetStr("TableName")
	If not BBS.Fun.isInteger(TableName) then
		GoBack"","������������������д��"
		Exit Sub
	End If
	If Int(TableName)=0 Then
		GoBack"","���ݱ�������Ϊ0"
		Exit Sub
	End If
	AllTable=Split(BBS.BBStable(0),",")
	For i=0 To uBound(AllTable)
	If Int(TableName)=Int(AllTable(i)) then
		GoBack"","���ݱ����Ѿ����ڣ�"
		Exit Sub
	End if
	Next
	S=BBS.BBStable(0)&","&TableName&"|"&BBS.BBStable(1)
	BBS.execute("update [config] Set BbStable='"&S&"'")
	BBS.execute("CREATE TABLE [bbs"&TableName&"](BbsID int IDENTITY (1, 1) NOT NULL CONSTRAINT PrimaryKey PRIMARY KEY,TopicID int Default 0,ReplyTopicID int Default 0,boardid int Default 0,Name varchar(20),Caption varchar(255),Content text,Face int Default 0,AddTime Datetime,LastTime datetime,IsDel byte Default 0,Ip varchar(40),IsAppraise byte Default 0,UbbString varchar(255))")
	BBS.execute("create index TopicID on [bbs"&TableName&"] (TopicID)")
	BBS.execute("create index boardid on [bbs"&TableName&"] (boardid)")
	BBS.execute("create index ReplyTopicID on [bbs"&TableName&"] (ReplyTopicID)")
	S="��� Bbs"&TableName&" ���ݱ�ɹ�!"
	BBS.NetLog"������̨_"&S
	Suc"",S,"admin_action.asp?action=SqlTable"
	BBS.Cache.clean("parameter")
End Sub

Sub DelSqlTable
	Dim ID,S,AllTable,I
	ID=request.querystring("ID")
	If int(ID)=int(BBS.BBStable(1)) Then
		GoBack "","�ñ��趨ΪĬ��ʹ�ñ�����ɾ����"
		Exit Sub
	End if
	AllTable=Split(BBS.BBStable(0),",")
	S=""
	For i=0 To uBound(AllTable)
		If int(ID)=Int(AllTable(i)) Then S="yes"
	Next
	If S="" Then
		Goback"ϵͳ����","��Ч�����ݱ����ƣ�":Exit Sub
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
	S="ɾ������Ϊ Bbs"&ID&" �����ݱ�("&S&"ƪ����)!"
	BBS.NetLog"������̨_"&S
	Suc"","�ɹ���"&S,"admin_action.asp?action=SqlTable"
End Sub

Sub SqlTableUnite
	Dim ID1,ID2,S,AllTable,i
	ID1=request.form("SqlTableID1")
	ID2=request.form("SqlTableID2")
	If ID1="0" or ID1="" or ID2="" or ID2="0" Then
	GoBack"","û��ѡ����"
	Exit Sub
	End If
	If ID1=ID2 Then
		GoBack "","ͬһ�����ݱ��úϲ�����~��"
		Exit Sub
	End If
	
	If int(ID1)=int(BBS.BBStable(1)) Then
		GoBack "","ָ�����ݱ���Ĭ��ʹ�ñ����ܺϲ���Ŀ���"
		Exit Sub
	End if
	'����
	AllTable=Split(BBS.BBStable(0),",")
	For i=0 To uBound(AllTable)
		If int(ID1)=Int(AllTable(i)) Then S=S&"Y"
		If Int(ID2)=Int(AllTable(i)) Then S=S&"Y"
	Next
	If S<>"YY" Then Goback"ϵͳ����","��Ч�����ݱ����ƣ�":Exit Sub
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
	S="���ݱ�BBS"&ID1&"�ϲ������ݱ�BBS"&ID2&"�ɹ�!"
	BBS.NetLog "������̨_"&S
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
	BBS.NetLog"������̨_��̳ϵͳ����ɹ�!"
	Suc "","��̳ϵͳ����ɹ�!�����<li>��������"&EssayNum&" | ��������"&TopicNum&" | ����������"&TodayNum&" | ע���û�����"&UserNum&" | ����ע���û���"&NewUser&"","admin_action.asp?action=UpdateBbs"
	BBS.Cache.clean("InfoUpdate")
End Sub

Sub UpdateTopic
	Dim Caption,Content,ID1,ID2,LastReply,Go,ReplyNum,Rs1,AllTable,i,S,MaxID,II
	ID1=BBS.Fun.GetStr("id1")
	ID2=BBS.Fun.GetStr("id2")
	MaxID=BBS.execute("select max(TopicID)from [Topic]")(0)
	IF ID1="" Then
		ID1=1:ID2=100
		Go= "��ʼ����"
	Else
		IF not isnumeric(ID1) or not isnumeric(ID2) then GoBack"","<li>����������д��":Exit Sub
		Set Rs=BBS.execute("Select TopicId,SqlTableID,Name From [Topic] where TopicID>="&ID1&" And Topicid<="&ID2&"")
		IF Not Rs.eof then
			AllTable=Split(BBS.BBStable(0),",")
			Do while not rs.eof
			For i=0 To uBound(AllTable)
				LastReply=Rs(2)&"|���޻ظ�"
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
	Go="��������"
	End If
	If Int(ID1)>Int(MaxID) Then
		Suc"�������","ȫ������ɹ�!","admin_action.asp?action=UpdateBbs"
		BBS.NetLog"������̨_�����޸���������"
		Exit Sub
	End If
	Caption="������������"
	Content="<form method='POST' style='margin:0' action='?action=UpdateTopic' onSubmit='ok.disabled=true;ok.value=""��������-���Եȡ�����""'>����д��Ҫ������������ӵĿ�ʼID�ͽ���ID��������֮�䲻Ҫ���̫��<br>�����̳������������ ID Ϊ��"&MaxID&"<br>��ʼID��<input type='text' name='ID1' size='20' value='"&ID1&"' class='text'><BR>����ID��<input type='text' name='ID2' size='20' value='"&ID2&"' class='text' /><BR><input type='submit' name='ok' value='"&Go&"' class='button' /><input type='reset' value='�� ��' class='button'></form>"
	ShowTable caption,Content
End Sub


Sub UpdateAllUser
	Dim Content,ID1,ID2,GoodNum,EssayNum,Rs1,Go,AllTable,I,S,MaxID,Flag
	ID1=BBS.Fun.GetStr("id1"):ID2=BBS.Fun.GetStr("id2")
	MaxID=BBS.execute("select max(id) from [User]")(0)
	IF ID1="" Then
		ID1=1:ID2=100
		Go= "��ʼ����"
	Else
		IF not isnumeric(ID1) or not isnumeric(ID2) then GoBack"","<li>����������д��":Exit Sub
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
	'���µȼ�
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
			ElseIf Flag=1 Then'���Ϊ������
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
	Go="��������"
	End If
	If Int(ID1)>Int(MaxID) Then
		BBS.NetLog"������̨_�����޸��û�����"
		Suc"�������","ȫ������ɹ�!","admin_action.asp?action=UpdateBbs"
		Exit Sub
	End If
	Content="<form method='POST' style='margin:0' action='?action=UpdateAllUser' onSubmit='ok.disabled=true;ok.value=""��������-���Եȡ�����""'>����д��Ҫ�����û��Ŀ�ʼID�ͽ���ID��������֮�䲻Ҫ���̫��<br />��̳ע���û����� ID Ϊ��"&MaxID&"<br />��ʼID��<input type='text' class='text' name='id1' size=20 value='"&ID1&"' /><br />����ID��<input type='text' class='text' name='id2' size='20' value='"&ID2&"' /><br /><input name='ok' class='button' type=submit value="&Go&" /><input type='reset' value='�� ��' class='button' /></form>"
	ShowTable "�û������޸�",Content
End Sub

Sub DelWuiong
	Dim i,AllTable,content
	Response.Write"<div class='mian'><div class='top'>��̳��������</div>"&_
	"<div class='divth'><b><span id='BBST'></span></b><div class='mian' style='margin:2px auto 0;width:400px;height:9'><img src='Images/icon/hr1.gif' width=0 height=16 id='BBSimg' align='absmiddle' alt='������' /></div>"&_
	"<div><span id='BBStxt' style='font-size:9pt'>0</span>%</div></div></div>"
	Response.Flush
	'BBS.execute("delete * from [admin] where (boardid<>0 and boardid<>-1) and (boardid not in(select boardid from [Board] where parentID<>0) or name not in(select name From [user] where isdel=1))")
	Call PicPro(0,8,"����������Ч���������Եȡ�����")	
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
		Show"������Ч������ϣ�"
			
	Call PicPro(1,8,"����������Ч���⣡���Եȡ�����")	
		AllTable=Split(BBS.BBStable(0),",")
		For i=0 To uBound(AllTable)
			BBS.execute("delete * from [bbs"&AllTable(i)&"] where TopicID<>0 and not exists (select name from [topic] where [bbs"&AllTable(i)&"].TopicId=[Topic].TopicID)")
			BBS.execute("delete * from [Topic] where SqlTableID="&AllTable(i)&" and not exists (select name from [bbs"&AllTable(i)&"] where [Topic].TopicID=[bbs"&AllTable(i)&"].TopicId)")
		Next
		Show"��Ч����������ϣ�"
	
	Call PicPro(2,8,"����������Ч��������¼")
		BBS.execute("delete * from [Appraise] where  not exists (select name from [Topic] where [Appraise].TopicID=[Topic].TopicId)")
		Show "��Ч������¼������ϣ�"	

	Call PicPro(3,8,"����������ЧͶƱ�����Եȡ�����")
		BBS.execute("delete * from [TopicVote] where  not exists (select name from [Topic] where [TopicVote].TopicID=[Topic].TopicId)")
		BBS.execute("delete * from [TopicVoteUser] where  not exists (select name from [Topic] where [TopicVoteUser].TopicID=[Topic].TopicId)")
		Show"��ЧͶƱ������ϣ�"
	
	Call PicPro(4,8,"����������Ч���ԣ����Եȡ�����")
		BBS.execute("delete * from [Sms] where not exists (select name from [User] where [Sms].MyName=[User].Name)")
		Show"��Ч����������ϣ�"
	Call PicPro(5,8,"����������Ч���棡���Եȡ�����")
		BBS.execute("delete * from [Placard] where not exists (select name from [User] where [Placard].Name=[User].Name)")
		If Not IsArray(BBS.Board_Rs) Then BBS.GetBoardCache()
		If IsArray(BBS.Board_Rs) Then
			For i=0 To Ubound(BBS.Board_Rs,2)
			'����ǰ��Ϊ��
			If BBS.Board_Rs(0,i)=0 Then
				BBS.execute("delete * from [Placard] where boardid<0 or boardid="&BBS.Board_Rs(1,i))
			End If
			Next
		End If
		Show"��Ч����������ϣ�"
	
	Call PicPro(6,8,"��������ɾ���û������ӣ����Եȡ�����")
		For i=0 To uBound(AllTable)
		BBS.execute("delete * from [bbs"&AllTable(i)&"] where not exists (select name from [User] where [bbs"&AllTable(i)&"].Name=[User].Name)")
		Next
		BBS.execute("delete * from [Topic] where not exists (select name from [User] where [Topic].Name=[User].Name)")
		Show "��Ч�û�������������ϣ�"
	
	Call PicPro(7,8,"����������Ч�ظ����ӣ�ʱ���Ƚϳ��������Եȡ�����")
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
	Show"��Ч�ظ�������ϣ�"		
	Response.Write "<script>document.getElementById(""BBSimg"").width=400;document.getElementById(""BBStxt"").innerHTML=""100"";BBST.innerHTML=""<font color=red>�ɹ�ȫ�������������</font>"";</script>"
	BBS.NetLog"������̨_������̳����"
End Sub
'������
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
		GoBack "","����������д!"
		Exit Sub
	End IF
	If Pass<>"1" Then Pass="0"
	If Ispic<>"1" Then Ispic="0"
	If IsIndex<>"1" Then IsIndex="0"
	BBS.Execute("Update [Link] Set Orders="&Orders&",Pass="&pass&",IsPic="&IsPic&",IsIndex="&IsIndex&" where ID="&ID&"")
	Next
	 SetLinkPage
	 BBS.NetLog"������̨_����������̳����"
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
	objname.Write"<!--#include file=""inc.asp""--><"&"% BBS.Head""linkinfo.asp"","""",""��վ��������"""&VbCrLf&"Call BBS.ShowTable(""<div>��վ��������</div>"",""<table width='100%'  border='0' align='center' cellpadding='0' cellspacing='5'><tr>"&TempText&"</tr></table><table width='100%'  border='0' align='center' cellpadding='5' cellspacing='0'><tr>"&TempPic&"</tr></table>"")"&VbCrLf&"BBS.Footer()"&VbCrLf&"Set BBS =Nothing%"&">"
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
		S="�����̳�������ӳɹ�!"
	Else
		S=BBS.execute("select Count(ID) From[Link]")(0)
		S=Int(S+1)
		BBS.execute("insert into [Link] (Bbsname,Url,Pic,Readme,admin,Orders,IsPic,pass)values('"&BbsName&"','"&Url&"','"&Pic&"','"&Readme&"','"&admin&"',"&S&","&IsPic&","&Pass&")")
		S="�޸���̳�������ӳɹ�!"
	End If
	BBS.NetLog "������̨_"&S
	Suc"",S,"admin_actionList.asp?action=Link"
	SetLinkPage
End Sub

Sub DelLink
	Dim ID,S
		ID=request.querystring("ID")
		BBS.execute("delete from [link] where ID="&ID&"")
		SetLinkPage
		S="ɾ����̳�������ӳɹ�!"
		BBS.NetLog"������̨_"&S
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
	Caption="ѹ�����ݿ�"
	Content="<b>ע�⣺</b>�������ݿ��������·���������������ݿ����ƣ��������ʹ�������ݿⲻ��ѹ������ѡ�񱸷����ݿ����ѹ��������<hr size=1>"&_
	"<form style='margin:0' method='post'>ѹ�����ݿ⣺<input type='text' name='DbPath' value='"&DbPath&"'>&nbsp;<input type='submit' class='button' value='��ʼѹ��' /><br><form>"&_
	"<input type='checkbox' name='boolIs97' value='True'>���ʹ�� Access 97 ���ݿ���ѡ��(Ĭ��Ϊ Access 2000 ���ݿ�)"
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
	BBS.NetLog"������̨_ѹ�����ݿ�"
	Suc "","������ݿ� " & request("Path") & "���Ѿ�ѹ���ɹ�!" ,"?action=CompressData"
End Sub
Sub NotCompactDB
	BBS.NetLog"������̨_ѹ�����ݿ�ʧ��!"
	GoBack "", "���ݿ����ƻ�·������ȷ������ѹ�����̷������⣡ �����ԣ�" 
End Sub


Sub BackupData()
Dim Caption,Content
Caption="������̳����"
Content="<b>ע�����</b><br>��̳���ݿⱸ�ݼ�����վ��ÿ��������£�<br>Ϊ��֤�������ݰ�ȫ������ʱ�벻Ҫ��Ĭ�������������������ݿ⡣<br>�������ݶ�ʧ��ʱ�򣬾Ϳ���������󱸷ݵ����ݿ�ָ���<br>ע�⣺����·��������������ռ��Ŀ¼�����·��<hr size=1>"&_
"<form style='margin:0' method='post' action='?action=BackupData&Go=Start'>��ǰ���ݿ�·��(���·��)��<input type=text size=15 name=DbPath value='data/db.mdb'><br>"&_
"�������ݿ�Ŀ¼(���·��)��<input type=text size='15' name='BkFolder' value='Data_Backup'>&nbsp;��Ŀ¼�����ڣ������Զ�����<BR>"&_
"�������ݿ�����(��д����)��<input type=text size=15 name=BkDbName value='Bak_db.mdb'>&nbsp;�籸��Ŀ¼�и��ļ��������ǣ���û�У����Զ�����<BR>"&_
"<input type='submit' value='��ʼ����' class='button'></form>"
ShowTable Caption,Content
If request("Go")="Start" then
	Dim fso,DbPath,BkFolder,BkDbName
	On error resume next
		DbPath=BBS.Fun.GetForm("DbPath")
		DbPath=server.mappath(DbPath)
		BkFolder=BBS.Fun.GetForm("BkFolder")
		BkDbName=BBS.Fun.GetForm("BkDbName")
		If Not IsAccess(Dbpath) Then 
			BBS.NetLog"������̨_�������ݿ�ʧ��!"
			GoBack"","���ݵ��ļ����ǺϷ������ݿ⡣"
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
			Caption="���ݳɹ�":Content="�������ݿ�ɹ�!�����ݵ����ݿ�·��Ϊ " &BkFolder& "\"& BkDbName
			BBS.NetLog"������̨_"&Content
		Else
			Caption="������Ϣ":Content="�Ҳ���������Ҫ���ݵ��ļ���"
			BBS.NetLog"������̨_�������ݿ�ʧ��!"
		End if
	ShowTable Caption,Content
End if
End sub
'---���ĳһĿ¼�Ƿ����-----
Function CheckDir(FolderPath)
Dim Fso1
	Folderpath=Server.MapPath(".")&"\"&folderpath
    Set fso1 = CreateObject("Scr"&"ipting"&".Fil"&"eSy"&"stemOb"&"ject")
    If fso1.FolderExists(FolderPath) then
       '����
       CheckDir = True
    Else
       '������
       CheckDir = False
    End if
    Set fso1 = nothing
End Function
'---����ָ����������Ŀ¼-----
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
Caption="�ָ���̳����"
Content="<b>ע�����</b>�ָ����ݿ� һ���������ָ�(���ݶ�ʧ���ƻ�)�ĵ�ǰʹ�����ݿ⡣<br>���ñ��ݵ����ݿ�ֱ�Ӱѵ�ǰʹ�õ����ݿ�ֱ�Ӹ��ǣ���ע�⣡<br>�����·��������������ռ��Ŀ¼�����·����<hr size=1 />"&_
"<form method='post' style='margin:0' action='?action=RestoreData&Go=Start'>�������ݿ�(���·��)��<input type='text' size='30' name='BackPath' value='Data_Backup\Bak_db.mdb'> ����д�����ָ��ı����ļ�<BR>"&_
"��ǰ���ݿ�(���·��)��<input type='text' size='30' name='DbPath' value='data/db.mdb'> ��д����ǰʹ�õ����ݿ�<br /><input onclick=""if(confirm('�˲������������ݿ⣡����\n��ȷ��Ҫ�ñ��ݵ����ݿ⸲�ǵ�Ȼʹ�õ����ݿ��𣡣�'))form.submit()"" type='button' class='button' value='�ָ�����'></form> "
ShowTable Caption,Content
If request("Go")="Start" then
 Caption="������Ϣ"
 Dim FSO,Dbpath,BackPath
 	DbPath=BBS.Fun.GetForm("DbPath")
	BackPath=BBS.Fun.GetForm("BackPath")
	if BackPath="" or DbPath="" then
		Content="���ȫ����д������"	
	'ElseIF Lcase(Dbpath)<>Lcase(Db) Then
		'Content="������Ĳ��ǵ�ǰʹ�����ݿ�ȫ��!"	
	Else
	On error resume next
		DbPath=server.mappath(DbPath)
		BackPath=server.mappath(BackPath)
		
		
		If Not IsAccess(BackPath) Then
			GoBack"",Content&" ���ݵ��ļ����ǺϷ������ݿ⡣"
			Exit Sub
		End If
		
		Set Fso=server.CreateObject("Scr"&"ipting"&".Fil"&"eSy"&"stemOb"&"ject")
		if fso.fileexists(DbPath) then  					
		On Error Resume Next
		fso.copyfile BackPath,DbPath
			If err.number=0 then
			Caption="�ָ��ɹ�":Content="�ɹ��Ļָ����ݿ⣡"
			BBS.NetLog"������̨_��"&BackPath&Content
			Else
			Content="����Ŀ¼�²������ı����ļ���"
			Err.clear
			End If
		else
		Content= "���ǵ�ǰʹ�õ����ݿ�ȫ��"
	End if
 End IF
ShowTable Caption,Content
End If
End sub

Sub AllUpdateGrade
	Dim ID,orders,GradeName,EssayNum,SqlEssayNum,PIC,Spic,Flag,i,S,UpdateUser,Grouping
	Grouping=Int(request.form("Grouping"))
	If Grouping=0 Then'������
		For i=1 to request.form("ID").count
			ID = Replace(request.form("ID")(i),"'","")
			GradeName = Replace(request.form("GradeName")(i),"'","")
			EssayNum = Replace(request.form("EssayNum")(i),"'","")
			If GradeName="" Then
				GoBack "","�ȼ����Ʊ�����д!"
				Exit Sub
			End IF		
			IF Not isnumeric(ID) or Not isnumeric(EssayNum)  Then
				GoBack "","��������������д!"
				Exit Sub
			End IF
			If EssayNum=0 Then S="OK"
		Next
		If S<>"OK" Then
			GoBack "","����ʧ�ܣ����������һ���ȼ��������Ϊ<font color=red>0</font>!"
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
		S="���������ȼ���(ͬʱ������"&I&"λ��Ա)"
	ElseIF Grouping=1 Then
		S="���ⶨ�Ƶȼ���"
	ElseIF Grouping=2 Then
		S="ϵͳ�̶��ȼ���"
	End If
	BBS.Cache.Clean("GradeInfo")
	 BBS.NetLog"������̨_����"&S&"�ɹ�!"
	 Suc"","����"&S&"�ɹ�!","admin_action.asp?action=Grade"
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
	S="ɾ���ȼ��� "&S&" �ɹ�!ͬʱ������"&I&"λ��Ա"
	BBS.NetLog"������̨_"&S
	Suc"","ɾ���ȼ��� "&S,"admin_action.asp?action=Grade"
End Sub

Sub SaveGrade
Dim S,i,ID,Strings,GradeName,EssayNum,Pic,Spic,Grouping,Flag
	GradeName = BBS.Fun.GetStr("GradeName")
	EssayNum = BBS.Fun.GetStr("EssayNum")
	Pic = BBS.Fun.GetStr("Pic")
	Spic=BBS.Fun.GetStr("Spic")
	ID=Request.form("ID")
	Grouping=Request.form("Grouping")
	If GradeName="" Then GoBack "","�ȼ����Ʊ�����д!":Exit Sub
	Strings=BBS.Fun.GetStr("S0")&"|"
	If len(S)>8 Then GoBack"�û�����ɫ��д����ȷ","":Exit Sub
	For i=1 to 37
		IF Request.form("S"&i)="" Then GoBack"","":Exit Sub
		If Not BBS.Fun.isInteger(Request.form("S"&i)) then
			GoBack "","һЩ����������Ϊ��������������̳�����������С�"
			Exit Sub
		End if
		Strings=Strings&Request("S"&i)&"|"
	Next
	Strings=Strings&"0|0|0"
	If Grouping=2 And ID<>"" Then
		BBS.execute("Update [Grade] Set GradeName='"&GradeName&"',PIC='"&PIC&"',Spic='"&Spic&"',Strings='"&Strings&"' where Grouping=2 and ID="&ID)
		S="�༭ϵͳ�̶��ȼ���ɹ�!"
		BBS.NetLog"������̨_"&S
		Suc"",S,"admin_action.asp?action=Grade"
	Else	
		If Grouping=0 Then
			IF Not BBS.Fun.isInteger(EssayNum)  Then GoBack "","��������������д!":Exit Sub
			Flag=0
			S="���������ȼ���"
		ElseIf Grouping=1 Then
			EssayNum=0
			Flag=1
			S="���ⶨ�Ƶȼ���"
		End If
		
		If ID<>"" Then
			BBS.execute("Update [Grade] Set Grouping="&Grouping&",GradeName='"&GradeName&"',EssayNum="&EssayNum&",PIC='"&PIC&"',Spic='"&Spic&"',Strings='"&Strings&"',Flag="&Flag&" where ID="&ID)
			S="�༭"&S&"�ɹ�!"
		Else
			BBS.execute("insert into [Grade] (Grouping,GradeName,EssayNum,PIC,Spic,Flag,Strings)values("&Grouping&",'"&GradeName&"',"&EssayNum&",'"&PIC&"','"&Spic&"',"&Flag&",'"&Strings&"')")
			S="���"&S&"�ɹ�!"
		End If
		BBS.Cache.Clean("GradeInfo")
		BBS.NetLog"������̨_"&S
		Suc"",S&"<li>ע�⣺�����û�Ҫ���´����µ�½�Ż���Ч</li>","admin_action.asp?action=Grade"
	End If
End Sub

Sub AllSms
	Dim SmsContent,UserType,Sql,Mrs,I
	SmsContent=BBS.Fun.GetStr("content")
	UserType=BBS.Fun.GetStr("caption")
	If SmsContent="" Then GoBack"","":Exit Sub
	If Len(SmsContent) >3000 Then GoBack"","�ַ�����":Exit Sub
If UserType="1" Then
	Dim Temp,OnlineCache,Eachonline
	OnlineCache=BBS.Cache.Value("OnlineCache")
	EachOnline=Split(OnlineCache,",")
	For I=0 to uBound(EachOnline)
	Temp=Split(EachOnline(I),"|")
	BBS.Execute("insert into [sms](name,MyName,Content,MyFlag) values('��̳С��ʹ','"&Temp(1)&"','"&SmsContent&"',1)")
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
	BBS.Execute("insert into [sms](name,MyName,Content,MyFlag) values('��̳С��ʹ','"&MRs(0,i)&"','"&SmsContent&"',1)")
	BBS.Execute("update [user] set NewSmsNum=NewSmsNum+1,SmsSize=SmsSize+1 where Name='"&MRs(0,i)&"'")
	BBS.UpdageOnline MRs(0,i),1
	Next
	End If
End IF
Suc"","�ɹ���Ⱥ�����ż�!","admin_SetHtmlEdit.asp?action=AllSms"
BBS.NetLog"������̨_Ⱥ���ż�"
End Sub


Sub Bank
	Dim Coin,UserType,Sql,Mrs,I,S,Flag
	UserType=BBS.Fun.GetStr("user")
	Coin=BBS.Fun.GetStr("Coin")
	Flag=BBS.Fun.GetStr("Flag")
	If UserType="" Then GoBack"","��ѡ���û�":Exit Sub
	If  Coin="" or Coin="0" Then GoBack"","":Exit Sub
	If Not isnumeric(Coin) Then GoBack"","����������д��":Exit Sub
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
If Flag=1 Then S="��Ǯ" Else S="��Ǯ"
Suc"","�ɹ���"&S&Coin&"!","admin_action.asp?action=bank"
BBS.NetLog"������̨_����"&Coin
End Sub

Sub AllBoardadmin
	Dim BoardadminName,Flag,boardid,Temp,S,i,GradeFlag
	BoardadminName=BBS.Fun.GetStr("Name")
	Flag=BBS.Fun.GetStr("Flag")
	If BoardadminName="" Then GoBack"","":Exit Sub
	Set Rs=BBS.execute("Select ID,Name,password,GradeFlag,EssayNum,IsVIP From[user] where name='"&BoardadminName&"'")
	If Rs.eof Then
		GoBack"","���ܲ��������û����ƻ�û��ע�ᡣ":Exit Sub
	End If
	If Flag="Add" Then
		If not BBS.Execute("select Name From[admin] where name='"&BoardadminName&"' and boardid<1").eof Then
			GoBack"","���û��Ѿ��ǳ���������վ���ˡ�":Exit Sub
		End If
			BBS.execute("Insert into[admin](name,[password],boardid)values('"&Rs(1)&"','"&Rs(2)&"',-1)")			
			BBS.UpdateGrade Rs(0),0,8
			S="�ɹ�������˳������� "&BoardadminName&" !"
	Else
		If BBS.Execute("select Name From[admin] where name='"&BoardadminName&"' and boardid=-1").eof Then GoBack"","���û����ǳ�������":Exit Sub
		BBS.Execute("Delete From[admin] where boardid=-1 And Name='"&BoardadminName&"'")
		If Not BBS.Execute("select Name From[admin]").eof Then
			GradeFlag=7
		Else
			GradeFlag=0
			If Rs(5)=1 Then GradeFlag=4
		End If
		BBS.UpdateGrade Rs(0),Rs(3),GradeFlag
		S="�ɹ������˳������� "&BoardadminName&" ��ְλ��"
	End if
	'����ˢ��
	BBS.UpdageOnline BoardadminName,3
	Rs.Close
	BBS.NetLog "������̨_"&S
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
		GoBack"","����ѡ��������̳���"
		Exit Sub
	End If
	Set Rs=BBS.execute("Select ID,Name,password,GradeFlag,EssayNum,IsVIP From[user] where name='"&BoardadminName&"'")
	If Rs.eof Then
	GoBack"","������Ӱ��������û����ƻ�û��ע�ᡣ":Exit Sub
	End If
	If Flag="Add" Then
		If Not BBS.Execute("select Name From[admin] where boardid="&boardid&" and Name='"&BoardadminName&"'").eof Then
			GoBack"","���û��Ѿ��Ǳ���İ����ˡ�":Exit Sub
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
			S="�ɹ�������˰��� "&BoardadminName&" !"
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
			If Rs(5)=1 Then'�����VIP
				BBS.UpdateGrade Rs(0),0,4
			Else
				BBS.UpdateGrade Rs(0),Rs(3),0
			End If
		End If
		S="�ɹ������˰��� "&BoardadminName&" ��ְλ��"
	End if
	Rs.Close
	'����ˢ��
	BBS.UpdageOnline BoardadminName,3
	BBS.NetLog "������̨_"&S
	BBS.Cache.clean("BoardInfo")
	Suc"",S,"admin_action.asp?action=Boardadmin"
End Sub

Sub DelFaction
	Dim Name
	Name=Request.QueryString("Name")
	BBS.Execute("Delete * From[Faction] where Name='"&Name&"'")
	BBS.Execute("update [User] Set Faction='' where Faction='"&Name&"'")
	BBS.NetLog "������̨_ɾ������ "&Name
	Suc"","�ɹ�ɾ���˰��ɣ�","admin_action.asp?action=Faction"
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
		GoBack"","�����˱�����ע���Ա��":Exit Sub
	End If
	If ID<>"" Then
		Set Rs=BBS.Execute("Select Name From[Faction] where ID="&ID)
		If Rs.Eof Then Goback"","��¼�ѱ�ɾ���ˣ�"
		If Name<>Rs(0) Then
		BBS.Execute("update [User] Set Faction='"&Name&"' where Faction='"&Rs(0)&"'")
		End If
		Rs.Close
		BBS.Execute("update [Faction] Set [Name]='"&Name&"',FullName='"&FullName&"',[Note]='"&Note&"',[User]='"&User&"',BuildDate='"&BuildDate&"' where ID="&ID)
		S="�ɹ��޸��˰��ɣ�"&Name
	Else
		BBS.execute("Insert into[Faction](Name,FullName,[Note],BuildDate,[User])Values('"&Name&"','"&FullName&"','"&Note&"','"&BuildDate&"','"&User&"')")
		BBS.Execute("update [User] Set Faction='"&Name&"' where Name='"&User&"'")
		S="�ɹ�����˰��ɣ�"&Name
	End If
		Suc"",S,"admin_action.asp?action=Faction"
	BBS.NetLog "������̨_"&S
End Sub

Sub DelEssay
	Dim UserName,DateNum,boardid,AllTable,I,SqlWhere,S
	DateNum=BBS.Fun.GetStr("DateNum")
	boardid=BBS.Fun.GetStr("boardid")
	UserName=BBS.Fun.GetStr("Name")
	AllTable=Split(BBS.BBStable(0),",")
	Select Case Request("Go")
	Case"Date"
		If not isnumeric(DateNum) Then GoBack"","����������������д��":Exit Sub
		If boardid=0 Then
			SQlwhere=""
			S="�ɹ�ɾ��"&DateNum&"��ǰ�������������������ظ�������"
		Else
			S="�ɹ�ɾ���� "&BBS.Execute("Select BoardName From[Board]where boardid="&boardid&"")(0)&" ��"&DateNum&"��ǰ�������������������ظ�������"
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
		Suc"",S&"<li>ɾ���������̳��һ��<a href=admin_action.asp?action=UpdateBbs>����</a>�����ؼư��������","admin_action.asp?action=DelEssay"
		BBS.NetLog"������̨_"&S	
	Case"DateNoRe"
		If not isnumeric(DateNum) Then GoBack"","����������������д��":Exit Sub
		If boardid=0 Then
			SQlwhere=""
			S="�ɹ�ɾ��"&DateNum&"��ǰû�лظ���������������������ظ�������"
		Else
			S="�ɹ�ɾ���� "&BBS.Execute("Select BoardName From[Board]where boardid="&boardid&"")(0)&" ��"&DateNum&"��ǰû�лظ���������������������ظ�������"
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
		Suc"",S&"<li>ɾ���������̳��һ��<a href=admin_action.asp?action=UpdateBbs>����</a>�����ؼư��������","admin_action.asp?action=DelEssay"
		BBS.NetLog"������̨_"&S	
	Case"User"
		If UserName="" Then GoBack"","":Exit Sub
		IF BBS.Execute("select name From[User] where Name='"&UserName&"'").eof Then
			GoBack"","����û����������ڣ�":Exit Sub
		End If
		If boardid=0 Then
			SQlwhere=""
			S="�ɹ�ɾ���û� "&UserName&" ���������ӣ���"
		Else
			S="�ɹ�ɾ���û� "&UserName&"�� "&BBS.Execute("Select BoardName From[Board]where boardid="&boardid&"")(0)&" �����ӣ�"
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
			Suc"",S&"<li>ɾ���������̳��һ��<a href=admin_action.asp?action=UpdateBbs>����</a>�����ؼư��������","admin_action.asp?action=DelEssay"
		BBS.NetLog"������̨_"&S
	Case Else
	GoBack"","�ύ��·������ȷ"
	End Select
End Sub

Sub DelSms
	Dim UserName,DateNum,boardid,S
	DateNum=BBS.Fun.GetStr("DateNum")
	Select Case Request("Go")
	Case"Date"
		If not isnumeric(DateNum) Then GoBack"","����������������д��":Exit Sub
		BBS.Execute("Delete From[Sms] where DATEDIFF('d',[AddTime],'"&BBS.NowBbsTime&"')>"&DateNum&"")
		S="�Ѿ��ɹ�ɾ����"&DateNum&"��ǰ�����������ż���"
	Case"User"
		UserName=BBS.Fun.GetStr("Name")
		IF UserName="" Then GoBack"","":Exit Sub
		IF BBS.Execute("select name From[User] where Name='"&UserName&"'").eof Then GoBack"","����û����������ڣ�":Exit Sub
		BBS.Execute("Delete From[Sms] where MyName='"&UserName&"'")
		S="�ɹ�ɾ�����û� "&UserName&" �����������ż���"
	Case"Auto"
		If not isnumeric(DateNum) Then GoBack"","����������������д��":Exit Sub
		BBS.Execute("Delete From[Sms] where DATEDIFF('d',[AddTime],'"&BBS.NowBbsTime&"')>"&DateNum&" And Name In ('�Զ�����ϵͳ','��̳С��ʹ')")
		S="�ɹ�ɾ����"&DateNum&"��ǰ��������̳�Զ����ŵ������ż���"
	End Select
	BBS.NetLog "������̨_"&S
	Suc"",S,"admin_action.asp?action=DelSms"
End Sub

Sub MoveEssay
	Dim boardid1,boardid2,DateNum,UserName,AllTable,I,S
	boardid1=BBS.Fun.GetStr("boardid1")
	boardid2=BBS.Fun.GetStr("boardid2")
	IF boardid1=boardid2 Then GoBack"","����û��ѡ��Ŀ����̳��":Exit Sub
	AllTable=Split(BBS.BBStable(0),",")
	DateNum=BBS.Fun.GetStr("DateNum")
	UserName=BBS.Fun.GetStr("Name")
Select Case Request("Go")
Case"Date"
	If not isnumeric(DateNum) Then GoBack"","����������������д��":Exit Sub
	Set Rs=BBS.Execute("Select TopicID,SqlTableID from[Topic] Where DATEDIFF('d',[AddTime],'"&BBS.NowBbsTime&"')>"&DateNum&"  and boardid="&boardid1) 
	Do while not Rs.eof
	BBS.Execute("update [Bbs"&Rs(1)&"] Set boardid="&boardid2&" where boardid="&boardid1&" and (TopicID="&Rs(0)&" or ReplyTopiciD="&Rs(0)&")")
	Rs.movenext
	Loop
	Rs.Close
	BBS.Execute("update [Topic] Set boardid="&boardid2&" where boardid="&boardid1&" And DATEDIFF('d',[AddTime],'"&BBS.NowBbsTime&"')>"&DateNum&"")
	S="�ɹ��İ�"&DateNum&"��ǰ�����Ӵ� "
Case"User"
	If UserName="" Then GoBack"","":Exit Sub
	IF BBS.Execute("select name From[User] where Name='"&UserName&"'").eof Then
		GoBack"","����û����������ڣ�":Exit Sub
	End IF
	Set Rs=BBS.Execute("Select TopicID,SqlTableID from[Topic] Where Name='"&UserName&"' And boardid="&boardid1) 
	Do while not Rs.eof
	BBS.Execute("update [Bbs"&Rs(1)&"] Set boardid="&boardid2&" where boardid="&boardid1&" and (TopicID="&Rs(0)&" or ReplyTopiciD="&Rs(0)&")")
	Rs.movenext
	Loop
	Rs.Close
	BBS.Execute("update [Topic] Set boardid="&boardid2&"  Where boardid="&boardid1&" and Name='"&UserName&"'")
	S="�ɹ��İ�"&UserName&"�����Ӵ� "
End Select
	S=S&BBS.Execute("select BoardName From[Board] where boardid="&boardid1&"")(0)&" �ƶ��� "&BBS.Execute("select BoardName From[Board] where boardid="&boardid2&"")(0)&"��"
	BBS.NetLog"������̨_"&S	
	Suc"",S&"�����ڽ���һ�� <a href='admin_Board.asp?action=BoardUpdate'>��������</a> ��","admin_action.asp?action=MoveEssay"
End Sub

Sub Clean
	Application.Contents.RemoveAll
	Suc "","���»���ɹ�","admin_action.asp?action=Clean"
	BBS.NetLog"������̨_���»���"
End Sub

Sub Topadmin
	Dim TopadminName,Flag,S,GradeFlag
	TopadminName=Replace(Request("Name"),"'","")
	Flag=Request("Flag")
	If TopadminName="" Then
		GoBack"","":Exit Sub
	End If
	Set Rs=BBS.execute("Select Name,password,ID,IsVip,EssayNum From[user] where name='"&TopadminName&"'")
	If Rs.eof Then GoBack"","���û����ƻ�û��ע�ᡣ":Exit Sub
	If Flag="1" Then
		If Not BBS.Execute("select Name From[admin] where boardid<1 and Name='"&TopadminName&"'").eof Then
			GoBack"","���û��Ѿ��ǹ���Ա��":Exit Sub
		End If
		BBS.execute("Insert into[admin](name,[password],boardid)values('"&Rs(0)&"','"&Rs(1)&"',0)")
		BBS.UpdateGrade Rs(2),0,9
		S="�ɹ������"
	Else
		BBS.Execute("delete * from [admin] where name='"&TopadminName&"' and boardid=0")
		GradeFlag=0
		If Rs(3)=1 Then Flag=4
		If Not BBS.Execute("select boardid from [admin] where name='"&TopadminName&"'").eof Then
			GradeFlag=7
		End if
		BBS.UpdateGrade Rs(2),Rs(4),GradeFlag
	 	S="�ɹ�������"
	End If
	Rs.Close
	S=S&BBS.GetGradeName(0,9)&":"&TopadminName&" !"
	BBS.UpdageOnline TopadminName,3
	BBS.NetLog"������̨_"&S
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
	BBS.NetLog"������̨_�޸�������"
	Response.redirect"admin_action.asp?action=GapAd"
End Sub

'�����Ƿ������ݿ�
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