<!--#include file="Inc.asp"-->
<%
Dim Caption,Content,Rs,ID,toUrl,Action,SetUserName,GoToUrl,Page,IsShow,Cue,ReplayNum,NotMe
If Not BBS.Founduser Then BBS.GoToerr(31)
BBS.CheckMake()
ID=BBS.Checknum(request.querystring("ID"))
Page=BBS.CheckNum(request.querystring("page"))
Action=lcase(request.querystring("Action"))
If Len(Action)>15 Then BBS.GotoErr(1)
If BBS.boardid=0 or ID=0  Then BBS.GoToErr(1)
Set Rs=BBS.Execute("Select Name,boardid,ReplyNum From [topic] Where TopicID="&ID&" And (boardid="&BBS.boardid&" or TopType=5 or TopType=4)")
IF Not Rs.eof Then
	SetUserName=Rs(0)
	BBS.boardid=Rs(1)
	ReplayNum=Rs(2)
	Rs.Close
Else
	BBS.GoToErr(58)
End IF
BBS.CheckBoard
If lcase(SetUserName)<>lcase(BBS.MyName) Then NotMe=True
If BBS.MyAdmin=7 Then
	If Not NotMe and action="ɾ��" Then
	Else
	'�Ǳ���İ��� 
	If Not BBS.IsBoardAdmin Then BBS.GotoErr(71)
	ENd If
End If
toURl="topic.asp?boardid="&BBS.boardid&"&id="&id&"&TB="&BBS.TB&"&Page="&Page
Cue=BBS.Row1("<div style=""padding:2px"">ע�⣺�����ʹ��������̳����Ȩ����ÿ�ι���Ͳ������ɶ��ᱻ��̳��־��¼��</div>")
If Action<>"�ƶ�" and Action<>"����" and Action<>"ɾ��" Then Affirm
GoToUrl=True
IsShow=True
BBS.Head "","","��������"
Caption="�� �� �� ����"
Select Case Action
Case"����"
	TopHeight
Case"����"
	SetTopicGood
Case"ȡ������"
	SetNotTopicGood
Case"�ö�"
	SetTop
Case"ȡ���ö�"
	SetNotTop
Case"���ö�"
	SetAllTop
Case"ȡ�����ö�"
	SetNotAllTop
Case"���ö�"
	SetClassTop
Case"ȡ�����ö�"
	SetNotClassTop
Case"����"
	SetTopicLock
Case"����"
	SetNotTopicLock
Case"ɾ��"
	Del
Case"�ƶ�"
	SetMove
Case"move"
	SaveMove
Case"�ѽ��"
	SetOk
Case"����"
	SetAppraise
Case"����"
	cover
Case"����"
	Setsubside
Case"saveappraise"
	SaveAppraise
Case"delappraise"
	delappraise
Case"editvote"
	EditVote
Case"savevote"
	SaveVote
Case Else
	BBS.GoToErr(1)
End Select
Set Rs=Nothing
If IsShow Then
	IF GoToUrl Then Content=Content&"<li><a href="&toUrl&">�ص�����</a></li>"
	Content="<div style=""margin:18px;line-height:150%"">"&Content&"</div>"
	BBS.ShowTable Caption,Content
End If
BBS.Footer()
Set BBS =Nothing

Sub SetTop
	If SESSION(CacheName& "MyGradeInfo")(31)="0" Then BBS.GotoErr(70)
		Set Rs=BBS.execute("Select TopType,caption From[Topic] where TopicID="&ID&" And boardid="&BBS.boardid&"")
		If Rs.eof Then
			BBS.GoToErr(58)
		Else
			IF Rs(0)=5 Then
				BBS.GoToErr(59)
			ElseIf Rs(0)=4 Then
				BBS.GoToErr(60)
			ElseIF Rs(0)=3 Then
				Caption="������Ϣ"
				Content="<li>�����������Ѿ����ö���</li>��"			
			Else
				BBS.Execute("update [Topic] Set TopType=3 where TopicID="&ID&" And boardid="&BBS.boardid&"")
				Content="<li>�趨Ϊ�ö�����---�ɹ���</li>"
				If NotMe Then
					BBS.execute("update [User] set Coin=Coin+"&Int(BBS.Info(96))&",Mark=Mark+"&Int(BBS.Info(97))&",GameCoin=GameCoin+"&Int(BBS.Info(98))&" Where name='"&SetUserName&"'")
					Content=Content&"<li>ͬʱ������������ߣ�"&SetUserName&" "&BBS.Info(120)&"+"&Int(BBS.Info(96))&" ��"&BBS.Info(121)&"+"&Int(BBS.Info(97))&"��"&BBS.Info(122)&"+"&Int(BBS.Info(98))&" �Ľ�����</li>"
				End If
				BBS.NetLog"������������ö���<br>����:"&left(Rs(1),20)&"<br>����:"&SetUserName
			End If
		End If
		Rs.Close
End Sub


Sub SetNotTop
	If SESSION(CacheName& "MyGradeInfo")(31)="0" Then BBS.GotoErr(70)
	Set Rs=BBS.execute("Select TopType,caption From[Topic] where TopicID="&ID&" And boardid="&BBS.boardid&"")
	If Rs.eof Then
		BBS.GoToErr(58)
	Else
		IF Rs(0)=5 Then
			BBS.GoToErr(59)
		ElseIf Rs(0)=4 Then
			BBS.GoToErr(60)
		ElseIF Rs(0)<>3 Then
			Caption="������Ϣ"
			Content="�����������Ѿ�û���ö��ˣ�"			
		Else
			BBS.Execute("update [Topic] Set TopType=0 where TopicId="&ID&" ")
			Content="<li>ȡ���ö�����---�ɹ���</li>"
			If NotMe Then
				BBS.execute("update [User] set Coin=Coin-"&Int(BBS.Info(96))&",Mark=Mark-"&Int(BBS.Info(97))&",GameCoin=GameCoin-"&Int(BBS.Info(98))&" Where Name='"&SetUserName&"'")
				Content=Content&"<li>ͬʱ������������ߣ�"&SetUserName&" "&BBS.Info(120)&" -"&Int(BBS.Info(96))&" ��"&BBS.Info(121)&" -"&Int(BBS.Info(97))&"��"&BBS.Info(122)&" -"&Int(BBS.Info(98))&"  �Ĳ�����</li>"	
			End If
			BBS.NetLog"�������ȡ���ö���<br>����:"&left(Rs(1),20)&"<br>����:"&SetUserName
		End If
	End If
	Rs.Close
End Sub
	
Sub SetAllTop
	If SESSION(CacheName& "MyGradeInfo")(33)="0" Then BBS.GotoErr(70)
	Set Rs=BBS.execute("Select TopType,caption From[Topic] where TopicID="&ID)
	If Rs.eof Then
		BBS.GoToErr(58)
	Else
		IF Rs(0)=5 Then
			BBS.GoToErr(59)
		Else
			BBS.execute("update [Topic] Set TopType=5 where TopicID="&id)
			Content="<li>�趨Ϊ���ö�����---�ɹ���</li>"
			If NotMe Then
			BBS.execute("update [user] Set Coin=Coin+"&Int(BBS.Info(90))&",Mark=Mark+"&Int(BBS.Info(91))&",GameCoin=GameCoin+"&Int(BBS.Info(92))&" where Name='"&SetUserName&"'")
			Content=Content&"<li>ͬʱ������������ߣ�"&SetUserName&" "&BBS.Info(120)&" +"&BBS.Info(90)&" ��"&BBS.Info(121)&" +"&BBS.Info(91)&"��"&BBS.Info(122)&" +"&BBS.Info(92)&" �Ľ�����</li>"
			End If
			BBS.NetLog"��������������ö���<br>����:"&left(Rs(1),20)&"<br>����:"&SetUserName
		End If
	End If
	Rs.Close
End Sub

Sub SetNotAllTop
	If SESSION(CacheName& "MyGradeInfo")(33)="0" Then BBS.GotoErr(70)
	Set Rs=BBS.execute("Select TopType,Caption From[Topic] where TopicID="&ID)
	If Rs.eof Then
		BBS.GoToErr(58)
	Else
		IF Rs(0)<>5 Then
			BBS.GoToErr(61)
		Else
		Content="<li>ȡ�����ö�����---�ɹ���</li>"
		BBS.execute("update [Topic] set TopType=0 where TopicID="&ID)
		If NotMe Then
			BBS.execute("update [user] set Coin=Coin-"&Int(BBS.Info(90))&",Mark=Mark-"&Int(BBS.Info(91))&",GameCoin=GameCoin-"&Int(BBS.Info(92))&" where name='"&SetUserName&"'")
			Content=Content&"<li>ͬʱ������������ߣ�"&SetUserName&" "&BBS.Info(120)&" -"&BBS.Info(90)&" ��"&BBS.Info(121)&" -"&BBS.Info(91)&"��"&BBS.Info(122)&" -"&BBS.Info(92)&" �Ĳ�����</li>"
		End If
		BBS.NetLog"�������ȡ�����ö���<br>����:"&left(Rs(1),20)&"<br>����:"&SetUserName
		End If
	End If
	Rs.Close
End Sub

Sub SetClassTop
	If SESSION(CacheName& "MyGradeInfo")(32)="0" Then BBS.GotoErr(70)
	Set Rs=BBS.execute("Select TopType,Caption From[Topic] where TopicID="&ID)
	If Rs.eof Then
		BBS.GoToErr(58)
	Else
		IF Rs(0)=5 Then
			BBS.GoToErr(59)
		ElseIf Rs(0)=4 Then
			BBS.GoToErr(60)
		Else
			BBS.execute("update [Topic] Set TopType=4 where TopicID="&id&" And boardid="&BBS.boardid&"")
			Content="<li>�趨Ϊ���ö�����---�ɹ���</li>"
			If NotMe Then
				BBS.execute("update [user] Set Coin=Coin+"&Int(BBS.Info(93))&",Mark=Mark+"&Int(BBS.Info(94))&",GameCoin=GameCoin+"&Int(BBS.Info(95))&" where Name='"&SetUserName&"'")
				Content=Content&"<li>ͬʱ������������ߣ�"&SetUserName&" "&BBS.Info(120)&" +"&BBS.Info(93)&" ��"&BBS.Info(121)&" +"&BBS.Info(94)&"��"&BBS.Info(122)&" +"&BBS.Info(95)&" �Ľ�����"
			End If
			BBS.NetLog"��������������ö���<br>����:"&left(Rs(1),20)&"<br>����:"&SetUserName
		End If
	End If
	Rs.Close
End Sub

Sub SetNotClassTop
	If SESSION(CacheName& "MyGradeInfo")(32)="0" Then BBS.GotoErr(70)
	Set Rs=BBS.execute("Select TopType,caption From[Topic] where TopicID="&ID)
	If Rs.eof Then
		BBS.GoToErr(58)
	Else
		IF Rs(0)<>4 Then
			BBS.GoToErr(62)
		Else
		BBS.execute("update [Topic] set TopType=0 where TopicID="&ID)
		Content="<li>ȡ�����ö�����---�ɹ���</li>"
		If NotMe Then
			BBS.execute("update [user] set Coin=Coin-"&Int(BBS.Info(93))&",Mark=Mark-"&Int(BBS.Info(94))&",GameCoin=GameCoin-"&Int(BBS.Info(95))&" where name='"&SetUserName&"'")
			Content=Content&"<li>ͬʱ������������ߣ�"&SetUserName&" "&BBS.Info(120)&" -"&BBS.Info(93)&" ��"&BBS.Info(121)&" -"&BBS.Info(94)&"��"&BBS.Info(122)&" -"&BBS.Info(95)&" �Ĳ�����</li>"
		End If
		BBS.NetLog"�������ȡ�����ö���<br>����:"&left(Rs(1),20)&"<br>����:"&SetUserName
		End If
	End If
	Rs.Close
End Sub

Sub SetTopicGood
	If SESSION(CacheName& "MyGradeInfo")(34)="0" Then BBS.GotoErr(70)
	Set Rs=BBS.Execute("Select IsGood,caption From[Topic] where TopicID="&ID&" And boardid="&BBS.boardid&"")
	If Rs.eof Then
		BBS.GoToErr(58)
	Else
		IF Rs(0)=1 Then
			Caption="������Ϣ"
			Content="�����������Ѿ��Ǿ��������ˣ�"
		ELse
			BBS.Execute("update [Topic] set IsGood=1 where TopicID="&ID&" And boardid="&BBS.boardid&"")
			BBS.execute("update [User] set GoodNum=GoodNum+1 where name='"&SetUserName&"'")
			Content="<li>�趨Ϊ��������---�ɹ���</li>"
		If NotMe Then
			BBS.execute("update [User] set Coin=Coin+"&Int(BBS.Info(99))&",Mark=Mark+"&Int(BBS.Info(100))&",GameCoin=GameCoin+"&Int(BBS.Info(101))&" where name='"&SetUserName&"'")
			Content=Content&"<li>ͬʱ������������ߣ�"&SetUserName&" "&BBS.Info(120)&" +"&BBS.Info(99)&" ��"&BBS.Info(121)&" +"&BBS.Info(100)&"��"&BBS.Info(122)&" +"&BBS.Info(101)&" �Ľ�����</li>"
		End If
			BBS.NetLog"����������þ�����<br>����:"&left(Rs(1),20)&"<br>����:"&SetUserName
		End If
	End If
	Rs.Close
End Sub

Sub cover
Dim S,Temp,BBSID
	If SESSION(CacheName& "MyGradeInfo")(27)="0" Then BBS.GotoErr(70)
	BbsID=BBS.CheckNum(request.querystring("BbsID"))
	Set Rs=BBS.execute("Select Isdel From [bbs"&BBS.TB&"]  where IsDel<>1 And BbsID="&BbsID&" And boardid="&BBS.boardid&"")
	If Rs.eof Then BBS.GotoErr(58)
	If Rs(0)=0 Then
		Temp=2
		S="��������"
	ELse
		Temp=0
		S="�����������"
	End If
	BBS.execute("update [bbs"&BBS.TB&"] set IsDel="&Temp&" where BbsID="&BbsID&" And boardid="&BBS.boardid&"")
	Content="<li>"&S&"---�ɹ���</li>"
	BBS.NetLog"�������ӣ�"&S
	Rs.close
End Sub

Sub SetNotTopicGood
	If SESSION(CacheName& "MyGradeInfo")(34)="0" Then BBS.GotoErr(70)
	Set Rs=BBS.Execute("Select IsGood,caption From[Topic] where TopicID="&ID)
	If Rs.eof Then
		BBS.GoToErr(58)
	Else
		IF Rs(0)=0 Then
			Caption="������Ϣ"
			Content="<li>�����������Ѿ���ȡ���˾����ˣ�</li>"
		ELse
			BBS.Execute("update [Topic] set IsGood=0 where TopicID="&ID)
			Content="<li>ȡ�����Ӿ���---�ɹ���</li>"
			If NotMe Then
				BBS.execute("update [User] set Coin=Coin-"&Int(BBS.Info(99))&",Mark=Mark-"&Int(BBS.Info(100))&",GameCoin=GameCoin-"&Int(BBS.Info(101))&",GoodNum=GoodNum-1 where name='"&SetUserName&"'")
				Content=Content&"<li>ͬʱ������������ߣ�"&SetUserName&" "&BBS.Info(120)&" -"&BBS.Info(99)&" ��"&BBS.Info(121)&" -"&BBS.Info(100)&"��"&BBS.Info(122)&" -"&BBS.Info(101)&" �Ĳ�����</li>"
			End If
			BBS.NetLog"�������ȡ��������<br>����:"&left(Rs(1),20)&"<br>����:"&SetUserName
		End If
	End If
	Rs.Close
End Sub

Sub SetTopicLock
	If SESSION(CacheName& "MyGradeInfo")(35)="0" Then BBS.GotoErr(70)
	Set Rs=BBS.execute("Select IsLock,caption From[Topic] where TopicID="&ID&" And boardid="&BBS.boardid&"")
	If Rs.eof Then
		BBS.GoToErr(58)
	Else
		IF Rs(0)=1 Then
			Caption="������Ϣ"
			Content="<li>�����������Ѿ��������ˣ�</li>"
		Else
			BBS.execute("update [Topic] set IsLock=1 where TopicID="&ID&" And boardid="&BBS.boardid&"")
			Content="<li>��������---�ɹ���</li>"
			BBS.NetLog"�����������������<br>����:"&left(Rs(1),20)&"<br>����:"&SetUserName
		End If
	End If
	Rs.Close
End Sub

Sub SetNotTopicLock
	If SESSION(CacheName& "MyGradeInfo")(35)="0" Then BBS.GotoErr(70)
	Set Rs=BBS.execute("Select Islock,caption From[Topic] where TopicID="&ID&" And boardid="&BBS.boardid&"")
	If Rs.eof Then
		BBS.GoToErr(58)
	Else
		IF Rs(0)=0 Then
			Caption="������Ϣ"
			Content="<li>�����������Ѿ������ˣ�</li>"
		Else
			BBS.execute("update [Topic] set IsLock=0 where TopicID="&ID&" And boardid="&BBS.boardid&"")
			Content="<li>���ӽ���---�ɹ���</li>"
			BBS.NetLog"����������������<br>����:"&left(Rs(1),20)&"<br>����:"&SetUserName
		End IF
	End if
	Rs.Close
End Sub
Sub DelMy(IsTopic)
	Dim BbsID
	BbsID=BBS.CheckNum(request.querystring("BbsID"))
	'ɾ���Լ�
	If IsTopic=1 Then
		BBS.execute("update [Topic] set IsDel=1 where TopicID="&ID)
		BBS.execute("update [bbs"&BBS.TB&"] set IsDel=1 where TopicID="&ID&" And boardid="&BBS.boardid&"")
		UpdateSys ReplayNum+1,1
	Else
		BBS.execute("update [bbs"&BBS.TB&"] set IsDel=1 where ReplyTopicID="&ID&" And BbsID="&BbsID&" And boardid="&BBS.boardid&"")
		Dim ReRs,TopicLastReply
		Set ReRs=BBS.execute("select top 1 Name,Content from [bbs"&BBS.TB&"] where boardid="&BBS.boardid&" And ReplyTopicID="&ID&" And IsDel=0 order by BbsID desc")
		If Not ReRs.Eof Then
				TopicLastReply=""&ReRs(0)&"|"&BBS.Fun.StrLeft(ReRs(1),40)
		Else
		        Dim RRs
		        Set RRs=BBS.execute("select Name from [Topic] where TopicId="&ID)
				If Not RRs.Eof Then
				   TopicLastReply=RRs(0)&"|���޻ظ�"
				Else
				   TopicLastReply="|���޻ظ�"
				End If
		        RRs.CLose:Set RRs=Nothing
		End If
		ReRs.CLose:Set ReRs=Nothing
		BBS.execute("Update [Topic] set ReplyNum=ReplyNum-1,LastReply='"&TopicLastReply&"' where TopicId="&ID&"")
		UpdateSys 1,0
	End If
	Caption="ɾ���ɹ�"
	Content="�Ѿ��ɹ�ɾ�������Լ���������ӣ�"
End Sub

Sub Del()
	Dim Temp,Cmd,Cause,IsSms,Sms,Smss,Mark,coin,GameCoin,S,BBSID,Sql
	with BBS
	If SESSION(CacheName& "MyGradeInfo")(26)="0" And SESSION(CacheName& "MyGradeInfo")(23)="0" Then .GotoErr(70)
	GotoUrl=False
	BbsID=.CheckNum(request.querystring("BbsID"))
	If BBSID<>0 Then
		Set Rs=.Execute("Select TopicID,name From [bbs"&.TB&"] where BbsID="&BbsID&" And boardid="&.boardid&"")
		IF Rs.eof Then .GoToErr(58)		
		If Rs(0)=ID Then'������
			BBSID=0
			If lcase(.MyName)=Lcase(Rs(1)) Then
				Call DelMy(1)
				Exit Sub
			End If
		Else
			If lcase(.MyName)=Lcase(Rs(1)) Then
				Call DelMy(0)
				Exit Sub
			End If
		End If
		Rs.close
	Else
	'������
		If lcase(.MyName)=Lcase(SetUserName) Then
			Call DelMy(1)
			Exit Sub
		End If
	End if
	Cmd=Request("Cmd")
	If SESSION(CacheName& "MyGradeInfo")(26)="0" Then .GotoErr(70)
	If Cmd="del" or .Info(51)="0" then'����ɾ��
		Affirm
		If .Info(51)="0" Then
			Coin=-.Info(108)
			Mark=-.Info(109)
			GameCoin=-.Info(110)
			Cause="-"
		Else
			Mark=.Fun.GetStr("mark")
			Coin=.Fun.GetStr("coin")
			GameCoin=.Fun.GetStr("gamecoin")
			Cause=.Fun.HtmlCode(.Fun.GetStr("cause"))
			IsSms=.Fun.GetStr("isSms")
			Sms=.Fun.GetStr("sms")
		End If
		If Cause="" Then
			Content="<li>����дɾ�����ɣ�<a href='javascript:history.go(-1)'>[����]</a></li>"	
		ElseIf Len(Cause)>10 Then
			Content="<li>ɾ�������������ܳ���10���ַ���<a href='javascript:history.go(-1)'>[����]</a></li>"	
		Else
			If BBSID<>0 Then
			Sql="Select IsDel,Name,Caption,ReplyTopicID From [bbs"&.TB&"] where BbsID="&BbsID&" And boardid="&.boardid&""
		Else
			Sql="Select IsDel,Name,Caption,ReplyNum From [Topic] where TopicID="&ID&" And boardid="&.boardid&""
		End If
		Set Rs=.execute(Sql)
			IF Rs.eof Then .GoToErr(58)
			IF Rs(0)<>1 Then
				GoToUrl=False
				If BBSID=0 Then
					.execute("update [Topic] set IsDel=1 where TopicID="&ID&" And boardid="&.boardid)
					.execute("update [bbs"&.TB&"] set IsDel=1 where TopicID="&ID&" And boardid="&.boardid&"")
					UpdateSys Rs(3)+1,1
				Else
					.execute("update [bbs"&.TB&"] set IsDel=1 where ReplyTopicID="&ID&" And BbsID="&BbsID&" And boardid="&.boardid&"")
					Dim ReRs,TopicLastReply
					Set ReRs=.execute("select top 1 Name,Content from [bbs"&.TB&"] where boardid="&.boardid&" And ReplyTopicID="&ID&" And IsDel=0 order by BbsID desc")
					If Not ReRs.Eof Then
							TopicLastReply=""&ReRs(0)&"|"&.Fun.StrLeft(ReRs(1),40)
					Else
		                    Dim RRs
		                    Set RRs=BBS.execute("select Name from [Topic] where TopicId="&ID&" And boardid="&.boardid)
		                    If Not RRs.Eof Then
		                        TopicLastReply=RRs(0)&"|���޻ظ�"
		                    Else
		                        TopicLastReply="|���޻ظ�"
		                    End If
		                    RRs.CLose:Set RRs=Nothing
					End If
					ReRs.CLose:Set ReRs=Nothing
					.execute("Update [Topic] set ReplyNum=ReplyNum-1,LastReply='"&TopicLastReply&"' where TopicId="&ID&"")
					.execute("update [bbs"&.TB&"] set IsDel=1 where TopicID="&ID&" And ReplyTopicID=0 And BbsID="&BbsID&" And boardid="&.boardid&"")
					UpdateSys 1,0
				End If
				Temp=GetGained(Rs(1),Coin,Mark,GameCoin)
			'����
				If IsSms="yes" Then
					Smss="�㷢������ӱ�ɾ����"&Cause&vbcrlf&Temp
					If Sms<>"" Then Smss=Smss&vbcrlf&vbcrlf&"�����ǲ����� "&.MyName&" ����ĸ���������Ϣ��"&vbcrlf&Sms
					.Execute("insert into [Sms](name,MyName,Content,MyFlag) values('�Զ�����ϵͳ','"&Rs(1)&"','"&Smss&"',1)")
					.Execute("update [User] set NewSmsNum=NewSmsNum+1,SmsSize=SmsSize+1 Where Name='"&Rs(1)&"'")
					.UpdageOnline Rs(1),1
				End If
				If Temp<>"" Then Temp="<li>"&Temp&"</li>"
				If BBSID=0 Then
					Content="<li>ɾ����������---�ɹ���</li>"&Temp
					.NetLog"ɾ��������:"&Cause&","&Temp&"<br>����:"&left(Rs(2),20)&"<br>����:"&SetUserName
				Else
					GotoUrl=True
					Content="<li>ɾ������---�ɹ���</li>"&Temp
					.NetLog"ɾ���ظ���:"&Cause&"<br>����:"&SetUserName
				End If
				Rs.Close
			Else
				Caption="������Ϣ"
				Content="<li>�����Ѿ�ɾ���ˣ�</li>"
			End IF
		End If
	Else
		IsShow=False
		If BBSID<>0 Then Caption="ɾ������" Else Caption="ɾ������"
		S="<form method=POST  style=""margin:0px"" action='?action=ɾ��&Cmd=del&TB="&.TB&"&id="&id&"&boardid="&.boardid&"&BBSID="&BBSID&"'>"
		S=S&Cue
		S=S&.Row("<b>�������ɣ�</b><select name='select' onChange='cause.value=this.options[this.selectedIndex].value'><option selected></option><option value='�����Ͻ����'>�����Ͻ����</option><option value='��������Υ��'>��������Υ��</option><option value='�������ĵ��ҹ�ˮ'>���ĵ��ҹ�ˮ</option><option value='�ظ���������'>�ظ���������</option></select>","<input name='cause' type='text' value='' size='30' maxlength='20'>������10���ַ�","65%","")
		S=S&.Row("<b>�ͷ�������</b>",.Info(120)&" <select name='coin'>"&Options(.Info(113),2)&"</select> "&.Info(121)&" <select name='mark'>"&Options(.Info(114),2)&"</select> "&.Info(122)&"<select name='gamecoin'>"&Options(.Info(115),2)&" </select>","65%","")
		S=S&.Row("<b>����֪ͨ�������ߣ�</b>","����<input name='issms' onclick='if(sms.disabled==true){sms.disabled=false;sms.focus()}else{sms.disabled=true;}' type='checkbox' value='yes'>&nbsp; ���Ը�����Ϣ��<input name='sms' size='30' type='text' value='' disabled='true'>","65%","")
		S=S&"<div style=""padding:2px;BACKGROUND: "&.SkinsPIC(2)&";"" align=""center""><input Class='button' type=""submit"" value=""ȷ������"" /></div></form>"
		.ShowTable Caption,S
	End If
	End with
End Sub

Sub SetMove
Dim S
	If SESSION(CacheName& "MyGradeInfo")(28)="0" Then BBS.GotoErr(70)
	IsShow=False
	S="<form method='POST' name='move' action='?action=move&TB="&BBS.TB&"&id="&id&"&boardid="&BBS.boardid&"' style=""margin:0px"" >"
	S=S&Cue
	S=S&BBS.Row("��ѡ������Ҫ�ƶ�������̳��",GetBoardList(),"65%","")
	If Lcase(SetUserName)<>Lcase(BBS.MyName) Then
	S=S&BBS.Row("�Ƿ�����֪ͨ�������ߣ�<input name='issms' onclick='if(sms.disabled==true){sms.disabled=false;sms.value=""֪ͨ���������ӱ�����Ա("&BBS.MyName&")�ƶ������""}else{sms.disabled=true;sms.value="""";}' type='checkbox' value='yes'>","<input name='sms' size='50' class='text' type='text' value='' disabled='true'>","65%","")
	End If
	S=S&"<div style=""padding:2px;BACKGROUND: "&BBS.SkinsPIC(2)&";"" align=""center""><input Class='button' type=""submit"" value=""ȷ������"" /></div></form>"
	BBS.ShowTable "�ƶ�����",S
End Sub

Sub SaveMove
	Dim IsSms,Sms,Newboardid,Temp
	with BBS
	If SESSION(CacheName& "MyGradeInfo")(28)="0" Then .GotoErr(70)
	GoToUrl=False
	IsSms=.Fun.GetStr("issms")
	Sms=.Fun.GetStr("sms")
	Newboardid=.Checknum(request.form("boardid"))
	If Newboardid=.boardid Then .GotoErr(62)
	.execute("update [Topic] Set boardid="&Newboardid&" where TopicID="&ID&"")
	.execute("update [bbs"&.TB&"] Set boardid="&Newboardid&" where TopicId="&ID&" or ReplyTopicid="&ID&"")
	If Lcase(SetUserName)<>Lcase(.MyName) Then
	If IsSms="yes" Then
		Sms=Sms&vbcrlf&"<a href=topic.asp?boardid="&Newboardid&"&id="&id&"&TB="&.TB&">����������������</a>"
		.Execute("insert into [Sms](name,MyName,Content,MyFlag) values('�Զ�����ϵͳ','"&SetUserName&"','"&Sms&"',1)")
		.Execute("update [User] set NewSmsNum=NewSmsNum+1,SmsSize=SmsSize+1 Where Name='"&SetUserName&"'")
		.UpdageOnline SetUserName,1
	End If
	End If
	'���°��
	Dim Boardupdate,LastReply
	Boardupdate=.GetEachBoardCache(.boardid)
	If Boardupdate(7)=ID&"" Then
	If .BoardString(6)=6 or .BoardString(6)=5 Then	'�������
	Else
		Set Rs=.execute("Select top 1 TopicID,Name,Caption,AddTime,Face,SqlTableID,boardid From [Topic] where IsDel=0 And boardid="&.boardid&" Order by LastTime Desc")
		If Not Rs.eof then
			LastReply=Rs("Name")&"|"&replace(.Fun.StrLeft(Rs("Caption"),22),"'","''")&"|"&Rs("AddTime")&"|"&Rs("Face")&"|"&Rs("TopicID")&"|"&Rs("boardid")&"|"&Rs("SqlTableID")
		End If
		Rs.Close
		Set Rs=Nothing
		.execute("Update [Board] set lastReply='"&LastReply&"' where boardid="&.boardid&" and Depth>0")
		Temp=.boardid
		If .BoardDepth>1 Then
			.Execute("Update [Board] set LastReply='"&LastReply&"' where boardid In ("&.BoardParentStr&") And Depth>0")
			Temp=Temp&","&.BoardParentStr
		End If
		.UpdateEcachBoardCache Temp,"0|0|0|"&LastReply
	End If
	End If
	Content="<li>�ƶ�����---�ɹ�����</li>"
	.NetLog"��������ƶ�"
	End with
End Sub

Function GetBoardList()
	Dim Temp,i
	Temp="<select Style='font-size: 9pt' name='boardid' >"
	If IsArray(BBS.Board_Rs) Then
		For i=0 To Ubound(BBS.Board_Rs,2)
		IF BBS.Board_Rs(0,I)=1 Then
			Temp=Temp&"<option value="&BBS.Board_Rs(1,I)&">��"&BBS.Board_Rs(3,I)&"</option>"
		ElseIf BBS.Board_Rs(0,I)=2 Then
			Temp=Temp&"<option value="&BBS.Board_Rs(1,I)&">�O��"&BBS.Board_Rs(3,I)&"</option>"
		End If
		Next
	End If
	GetBoardList=Temp&"</select>"
End Function


Sub UpdateSys(EssayNum,TopicNum)
	with BBS
	Dim LastReply,TempContent,TempID,Rs1
	.execute("update [Config] set AllEssayNum=AllEssayNum-"&EssayNum&",TopicNum=TopicNum-"&TopicNum)	
	'�����������治��ʾ���ظ�
	If .BoardString(6)=6 or .BoardString(6)=5 Then
		LastReply=""
	Else
		Set Rs1=.execute("Select top 1 TopicID,Name,Caption,AddTime,Face,SqlTableID,boardid From [Topic] where IsDel=0 And boardid="&.boardid&" Order by LastTime Desc")
		if Rs1.eof then
			LastReply=""
		Else
			LastReply=Rs1("Name")&"|"&replace(.Fun.StrLeft(Rs1("Caption"),22),"'","''")&"|"&Rs1("AddTime")&"|"&Rs1("Face")&"|"&Rs1("TopicID")&"|"&Rs1("boardid")&"|"&Rs1("SqlTableID")
		End If
		Rs1.Close
		Set Rs1=Nothing
	End If
	TempID=.boardid
	.Execute("update [Board] set LastReply='"&LastReply&"',EssayNum=EssayNum-"&EssayNum&",TopicNum=TopicNum-"&TopicNum&" where boardid="&.boardid&" and ParentID<>0")
	'���¸����
	If .BoardDepth>1 Then
		.Execute("Update [Board] set LastReply='"&LastReply&"',EssayNum=EssayNum-"&EssayNum&",TopicNum=TopicNum-"&TopicNum&" where boardid In ("&.BoardParentStr&") And ParentID<>0")
		TempID=TempID&","&.BoardParentStr
	End If
	.UpdateEcachBoardCache TempID,-EssayNum&"|"&-TopicNum&"|0|"&LastReply
	'����ϵͳ��̬��������
	TempContent=.InfoUpdate(0)-Int(EssayNum)&","&.InfoUpdate(1)-Int(TopicNum)&","&.InfoUpdate(2)&","&.InfoUpdate(3)&","&.InfoUpdate(4)&","&.InfoUpdate(5)&","&.InfoUpdate(6)&","&.InfoUpdate(7)&","&.InfoUpdate(8)&","&.InfoUpdate(9)&","&.InfoUpdate(10)
	.Cache.Add "InfoUpdate",TempContent,dateadd("n",2000,BBS.NowBBSTime)
	End with
End Sub

Sub TopHeight
	If SESSION(CacheName& "MyGradeInfo")(29)="0" Then BBS.GotoErr(70)
	BBS.Execute("update [Topic] set LastTime='"&BBS.NowBbsTime&"' where TopicID="&ID&" And boardid="&BBS.boardid&"")
	BBS.Execute("update [bbs"&BBS.TB&"] set LastTime='"&BBS.NowBbsTime&"' where TopicID="&ID&" And boardid="&BBS.boardid&"")
	Content="<Li>������������---�ɹ�����"
	BBS.NetLog"�����������"
End Sub

Sub Setsubside
	If SESSION(CacheName& "MyGradeInfo")(30)="0" Then BBS.GotoErr(70)
	BBS.Execute("update [Topic] set LastTime=LastTime-30 where TopicID="&ID&" And boardid="&BBS.boardid&"")
	Content="<Li>�Ѿ��ɹ���ʹ����������׵�һ����ǰ�������棡"
	BBS.NetLog"�����������"
End Sub

Function Options(Num,Flag)
	dim I,Steps,Num1,Num2
	Num1=-Num
	Num2=Num
	If Flag=1 Then Num1=0 
	If Flag=2 Then Num2=0
	Steps=1
	If Num>20 Then Steps=Num\10
	For I=Num1 to Num2 Step Steps
	Options=Options&"<option value="&I
	If I=0 Then Options=Options&" selected"
	Options=Options&">"&I&"</option>"
	Next
End Function

Sub delappraise
	Dim BbsID,S
	S="ɾ��������¼ "
	If SESSION(CacheName& "MyGradeInfo")(41)="0" Then BBS.GotoErr(70)
	BbsID=BBS.CheckNum(request.querystring("BbsID"))
	BBS.Execute("delete from [appraise] where BbsID="&BbsID&" and TopicID="&ID)
	BBS.Execute("update [bbs"&BBS.TB&"] set IsAppraise=0 where BbsID="&BbsID&" And boardid="&BBS.boardid&"")
	Content="<li>"&S&"---�ɹ���</li>"
	BBS.NetLog S
End Sub

Sub SetAppraise
	Dim BbsID,S
	If SESSION(CacheName& "MyGradeInfo")(36)="0" Then BBS.GotoErr(70)
	S=BBS.Execute("Select Count(*) From[Appraise] where AdminName='"&BBS.MyName&"' And DATEDIFF('d',[AddTime],'"&BBS.NowBbsTime&"')<1")(0)
	If IsNull(S) Then S=0
	If S>Int(BBS.Info(49)) Then BBS.GotoErr(66)
	IsShow=False
	BbsID=BBS.CheckNum(request.querystring("BbsID"))
	Set Rs=BBS.execute("Select BbsID From [bbs"&BBS.TB&"] where BbsID="&BbsID&" And boardid="&BBS.boardid&"")
	IF Rs.eof Then
		BBS.GoToErr(58)
	Else
		S="<form method=POST style=""margin:0px"" action='?action=saveappraise&TB="&BBS.TB&"&BbsID="&BbsID&"&id="&id&"&boardid="&BBS.boardid&"&Page="&Page&"'>"
		S=S&Cue
		S=S&BBS.Row("<b>�������ɣ�</b><select name='select' onChange='cause.value=this.options[this.selectedIndex].value'><option selected>���������Զ���</option><option value='���������Ӳ���Ŷ'>���������Ӳ���Ŷ</option><option value='������л��˽����'>������л��˽����</option><option value='���������¸�����'>���������¸�����</option><option value='���������Ͻ����'>���������Ͻ����</option><option value='������������Υ��'>������������Υ��</option><option value='�������ĵ��ҹ�ˮ'>�������ĵ��ҹ�ˮ</option><option value='�����ظ���������'>�����ظ���������</option></select>","<input name='cause' type='text' value='' size='30' maxlength='25'>������22���ַ�","65%","")
		S=S&BBS.Row("<b>����������</b>",BBS.Info(120)&" <select name='coin'>"&Options(BBS.Info(113),0)&"</select> "&BBS.Info(121)&" <select name='mark'>"&Options(BBS.Info(114),0)&"</select> "&BBS.Info(122)&"<select name='gamecoin'>"&Options(BBS.Info(115),0)&" </select>","65%","")
		S=S&BBS.Row("<b>����֪ͨ�������ߣ�</b>","����<input name='issms' onclick='if(sms.disabled==true){sms.disabled=false;sms.focus()}else{sms.disabled=true;}' type='checkbox' value='yes'>&nbsp; ���Ը�����Ϣ��<input name='sms' size='30' type='text' value='' disabled='true'>","65%","")
		S=S&"<div style=""padding:2px;BACKGROUND: "&BBS.SkinsPIC(2)&";"" align=""center""><input Class='button' type=""submit"" value=""ȷ������"" /></div></form>"
	End If
	BBS.ShowTable"��������",S
End Sub
Sub SaveAppraise
	If SESSION(CacheName& "MyGradeInfo")(36)="0" Then BBS.GotoErr(70)
	Dim BbsID,Cause,Mark,Coin,GameCoin,IsSms,Sms,Smss,temp
	BbsID=BBS.CheckNum(request.querystring("BbsID"))
	Cause=BBS.Fun.HtmlCode(BBS.Fun.GetStr("cause"))
	Mark=BBS.Fun.GetStr("mark")
	Coin=BBS.Fun.GetStr("coin")
	GameCoin=BBS.Fun.GetStr("gamecoin")
	IsSms=BBS.Fun.GetStr("issms")
	Sms=BBS.Fun.GetStr("sms")
	Caption="��������"
	Set Rs=BBS.execute("Select Name,Caption From [bbs"&BBS.TB&"] where BbsID="&BbsID&" And boardid="&BBS.boardid&"")
	IF Rs.eof Then
		BBS.GoToErr(58)
	ElseIf Lcase(Rs(0))=Lcase(BBS.MyName) Then
		Content="<li>���ܶ��Լ�����������</li>"
	ElseIf Cause="" And (Mark=0 and Coin=0 and GameCoin=0) Then
		Content="<li>����д�������ύ��</li>"	
	ElseIf Len(Cause)>22 Then
		Content="<li>���������������ܳ���25���ַ���</li>"	
	Else
		Cause=BBS.Fun.HtmlCode(Cause)
		BBS.execute("insert into [Appraise](BbsID,TopicID,Cause,Mark,Coin,GameCoin,AdminName,AddTime)VALUES("&BbsID&","&ID&",'"&Cause&"',"&Mark&","&Coin&","&GameCoin&",'"&BBS.MyName&"','"&BBS.NowBbsTime&"')")
		Temp=GetGained(Rs(0),Coin,Mark,GameCoin)
		If IsSms="yes" Then
			Smss="������ӣ�<a href="""&toUrl&""">"&Rs(1)&"</a><br>�����ۣ�"&Cause&"<br>"&Temp
			If Sms<>"" Then Smss=Smss&"<br><br>�����ǲ�����:"&BBS.MyName&" ����ĸ���������Ϣ��"&vbcrlf&Sms
			BBS.Execute("insert into [Sms](name,MyName,Content,MyFlag) values('�Զ�����ϵͳ','"&Rs(0)&"','"&Smss&"',1)")
			BBS.Execute("update [User] set NewSmsNum=NewSmsNum+1,SmsSize=SmsSize+1 Where Name='"&Rs(0)&"'")
			BBS.UpdageOnline Rs(0),1
		End If
		BBS.NetLog"��������:"&Cause&","&Temp
		Rs.Close
		BBS.Execute("Update [bbs"&BBS.TB&"] Set IsAppraise=1 where BbsID="&BbsID&" And boardid="&BBS.boardid&"")
		Caption="��������"
		If Temp<>"" Then Temp="<li>"&Temp&"</li>"
		Content="<li>���������ɹ�!</li>"&Temp
	End If
End Sub

Function GetGained(UserName,Coin,Mark,GameCoin)
	If Coin<>0 or Mark<>0 or GameCoin<>0 Then 
		GetGained="���Ҷ����� "&UserName&" ������"
		If Coin<>0 Then GetGained=GetGained&BBS.Info(120)&Coin&","
		If Mark<>0 Then GetGained=GetGained&BBS.Info(121)&Mark&","
		If GameCoin<>0 Then GetGained=GetGained&BBS.Info(122)&GameCoin
		GetGained=GetGained&"�Ĳ�����"
		BBS.Execute("Update [user] set Mark=Mark+"&Mark&",Coin=Coin+"&Coin&",GameCoin=GameCoin+"&GameCoin&" where Name='"&UserName&"'")
		If lcase(UserName)=Lcase(BBS.MyName) Then 	Session(CacheName & "MyInfo") = Empty
	Else
		GetGained=""
	End If
End Function

Sub EditVote()
	If SESSION(CacheName& "MyGradeInfo")(38)="0" Then BBS.GotoErr(70)
	IsShow=False
	Dim S,Vote,VoteType,VoteNum,i,II
	Set Rs=BBS.execute("Select Top 1 TopicID,Vote,VoteNum,VoteType,OutTime from [TopicVote] where TopicID="&ID)
	IF Not rs.eof Then
		Caption="ͶƱ���ӵı��⣺"&BBS.execute("Select Caption from [Topic] where TopicID="&ID)(0)
		Vote=split(Rs(1),"|")
		VoteNum=Split(Rs(2),"|")
		II=UBound(Vote)
		For I = 1 To II
			S=S&BBS.Row("&nbsp;"&i,"<input size='80' name='Votes"&i&"' type='text' value='"&Vote(i)&"'> ͶƱ����<input size='3' name='VoteNum"&i&"' type='text' value='"&VoteNum(i)&"'>","95%","")
		Next
		S=S&BBS.Row("&nbsp;�½�","<input size='80' name='Votes"&II+1&"' type='text' value=''> ͶƱ����<input size='3' name='VoteNum"&II+1&"' type='text' class='text' value='0'>","95%","")
		If Rs(3)=2 then VoteType=" checked"
		S="<form style=""margin:0px"" action='?action=SaveVote&TB="&BBS.TB&"&boardid="&BBS.boardid&"&id="&id&"' method='post'><input name='AutoValue' type='hidden' value='"&UBound(Vote)+1&"'>"&S&"<div style="" padding:5px;BACKGROUND: "&BBS.SkinsPIC(2)&";"" align=""center"">�Ƿ��ѡ��<input name='VoteType' type='checkbox' value=2 "&voteType&"> &nbsp; ����ʱ�䣺<input size=20 name='OutTime' type='text' value='"&Rs(4)&"'> &nbsp; <input type='submit' value='ȷ���޸�'></div></form>"
		BBS.ShowTable Caption,S
	Else
		BBS.GotoErr(58)
	End IF
	Rs.Close
End Sub

Sub SaveVote()
	Dim VoteValue,VoteType,Votes,VoteNum,OutTime,i,Temp
	VoteType=BBS.CheckNum(request.Form("VoteType"))
	If VoteType<>2 Then VoteType=1
	VoteValue=BBS.CheckNum(request.Form("AutoValue"))
	OutTime=BBS.Fun.GetStr("OutTime")
	If Not IsDate(OutTime) Then OutTime="2030-10-1 07:30:00"
	For i=1 to VoteValue
	Temp=Left(BBS.Fun.Checkbad(Trim(BBS.Fun.GetStr("Votes"&I))),250)
	IF Temp>"" Then
		Votes=Votes&"|"&Temp
		If Not BBS.Fun.isInteger(BBS.Fun.GetStr("VoteNum"&I)) Then BBS.GoToErr(61)
		VoteNum=VoteNum&"|"&BBS.Fun.GetStr("VoteNum"&I)
	End If
	Next
	If Votes<>"" Then
		BBS.Execute("Update [TopicVote] Set VoteType="&VoteType&",Vote='"&Votes&"',VoteNum='"&VoteNum&"',OutTime='"&OutTime&"' where TopicID="&ID)
		Temp="<li><a href='Topic.asp?TB="&BBS.TB&"&boardid="&BBS.boardid&"&id="&id&"'>�ص�ͶƱ����</a></li><li><a href='?action=EditVote&TB="&BBS.TB&"&boardid="&BBS.boardid&"&id="&id&"'>�����޸�ͶƱ����</a></li>"
	Else
		BBS.Execute("Delete From[TopicVote] where TopicID="&ID)
		BBS.Execute("Delete From[TopicVoteUser] where TopicID="&ID)
		BBS.Execute("Update [Topic] Set IsVote=False where TopicID="&ID)
		Temp="<li>�Ѿ��ɹ�ɾ����ͶƱ���ݣ�<li><a href='topic.asp?TB="&BBS.TB&"&boardid="&BBS.boardid&"&id="&id&"'>�ص���������</a></li>"
	End IF
	BBS.NetLog"��������:�޸�ͶƱѡ��"
	Caption="�����ɹ�"
	Content=Temp
End Sub

Sub SetOk
    Dim S_OK
	S_OK = 0
	Set Rs=BBS.Execute("Select TopicID,name From [Topic] where TopicID="&ID&" and boardid="&BBS.boardid)
	IF Rs.eof Then
	  BBS.GoToErr(58)
	Else
	  If lcase(BBS.MyName)=Lcase(Rs(1)) Then S_OK = 1
	End If		
	Rs.close
	If BBS.MyAdmin<7 and S_OK = 0 Then BBS.GotoErr(70)
	BBS.execute("update [Topic] set Caption='���ѽ����'&Caption where TopicID="&ID&" and boardid="&BBS.boardid&" and Caption not like'%���ѽ����%'")
	Content="<li>�趨����Ϊ�ѽ������---�ɹ�����</li>"
	BBS.NetLog"�趨����Ϊ�ѽ��"
End Sub

Sub Affirm()
If Request.ServerVariables("request_method") <> "POST" then
Response.write "<form name='KK' method=post action=?"&Request.ServerVariables("Query_String")&"></form><SCRIPT>if(confirm('��ȷ��Ҫִ�иò���?')){returnValue=KK.submit()}else{returnValue=history.back()}</SCRIPT>"
Response.End
End If
End Sub
%>