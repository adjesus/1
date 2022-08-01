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
	If Not NotMe and action="删除" Then
	Else
	'非本版的版主 
	If Not BBS.IsBoardAdmin Then BBS.GotoErr(71)
	ENd If
End If
toURl="topic.asp?boardid="&BBS.boardid&"&id="&id&"&TB="&BBS.TB&"&Page="&Page
Cue=BBS.Row1("<div style=""padding:2px"">注意：请谨慎使用您的论坛管理权力，每次管理和操作理由都会被论坛日志记录！</div>")
If Action<>"移动" and Action<>"评帖" and Action<>"删除" Then Affirm
GoToUrl=True
IsShow=True
BBS.Head "","","管理帖子"
Caption="操 作 成 功！"
Select Case Action
Case"提升"
	TopHeight
Case"精华"
	SetTopicGood
Case"取消精华"
	SetNotTopicGood
Case"置顶"
	SetTop
Case"取消置顶"
	SetNotTop
Case"总置顶"
	SetAllTop
Case"取消总置顶"
	SetNotAllTop
Case"区置顶"
	SetClassTop
Case"取消区置顶"
	SetNotClassTop
Case"锁定"
	SetTopicLock
Case"解锁"
	SetNotTopicLock
Case"删除"
	Del
Case"移动"
	SetMove
Case"move"
	SaveMove
Case"已解决"
	SetOk
Case"评帖"
	SetAppraise
Case"屏蔽"
	cover
Case"沉底"
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
	IF GoToUrl Then Content=Content&"<li><a href="&toUrl&">回到帖子</a></li>"
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
				Caption="错误信息"
				Content="<li>该主题帖子已经被置顶了</li>！"			
			Else
				BBS.Execute("update [Topic] Set TopType=3 where TopicID="&ID&" And boardid="&BBS.boardid&"")
				Content="<li>设定为置顶帖子---成功！</li>"
				If NotMe Then
					BBS.execute("update [User] set Coin=Coin+"&Int(BBS.Info(96))&",Mark=Mark+"&Int(BBS.Info(97))&",GameCoin=GameCoin+"&Int(BBS.Info(98))&" Where name='"&SetUserName&"'")
					Content=Content&"<li>同时给该主题的作者："&SetUserName&" "&BBS.Info(120)&"+"&Int(BBS.Info(96))&" ，"&BBS.Info(121)&"+"&Int(BBS.Info(97))&"，"&BBS.Info(122)&"+"&Int(BBS.Info(98))&" 的奖励！</li>"
				End If
				BBS.NetLog"主题管理：设置置顶。<br>主题:"&left(Rs(1),20)&"<br>作者:"&SetUserName
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
			Caption="错误信息"
			Content="该主题帖子已经没有置顶了！"			
		Else
			BBS.Execute("update [Topic] Set TopType=0 where TopicId="&ID&" ")
			Content="<li>取消置顶帖子---成功！</li>"
			If NotMe Then
				BBS.execute("update [User] set Coin=Coin-"&Int(BBS.Info(96))&",Mark=Mark-"&Int(BBS.Info(97))&",GameCoin=GameCoin-"&Int(BBS.Info(98))&" Where Name='"&SetUserName&"'")
				Content=Content&"<li>同时给该主题的作者："&SetUserName&" "&BBS.Info(120)&" -"&Int(BBS.Info(96))&" ，"&BBS.Info(121)&" -"&Int(BBS.Info(97))&"，"&BBS.Info(122)&" -"&Int(BBS.Info(98))&"  的操作！</li>"	
			End If
			BBS.NetLog"主题管理：取消置顶。<br>主题:"&left(Rs(1),20)&"<br>作者:"&SetUserName
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
			Content="<li>设定为总置顶帖子---成功！</li>"
			If NotMe Then
			BBS.execute("update [user] Set Coin=Coin+"&Int(BBS.Info(90))&",Mark=Mark+"&Int(BBS.Info(91))&",GameCoin=GameCoin+"&Int(BBS.Info(92))&" where Name='"&SetUserName&"'")
			Content=Content&"<li>同时给该主题的作者："&SetUserName&" "&BBS.Info(120)&" +"&BBS.Info(90)&" ，"&BBS.Info(121)&" +"&BBS.Info(91)&"，"&BBS.Info(122)&" +"&BBS.Info(92)&" 的奖励！</li>"
			End If
			BBS.NetLog"主题管理：设置总置顶。<br>主题:"&left(Rs(1),20)&"<br>作者:"&SetUserName
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
		Content="<li>取消总置顶帖子---成功！</li>"
		BBS.execute("update [Topic] set TopType=0 where TopicID="&ID)
		If NotMe Then
			BBS.execute("update [user] set Coin=Coin-"&Int(BBS.Info(90))&",Mark=Mark-"&Int(BBS.Info(91))&",GameCoin=GameCoin-"&Int(BBS.Info(92))&" where name='"&SetUserName&"'")
			Content=Content&"<li>同时给该主题的作者："&SetUserName&" "&BBS.Info(120)&" -"&BBS.Info(90)&" ，"&BBS.Info(121)&" -"&BBS.Info(91)&"，"&BBS.Info(122)&" -"&BBS.Info(92)&" 的操作！</li>"
		End If
		BBS.NetLog"主题管理：取消总置顶。<br>主题:"&left(Rs(1),20)&"<br>作者:"&SetUserName
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
			Content="<li>设定为区置顶帖子---成功！</li>"
			If NotMe Then
				BBS.execute("update [user] Set Coin=Coin+"&Int(BBS.Info(93))&",Mark=Mark+"&Int(BBS.Info(94))&",GameCoin=GameCoin+"&Int(BBS.Info(95))&" where Name='"&SetUserName&"'")
				Content=Content&"<li>同时给该主题的作者："&SetUserName&" "&BBS.Info(120)&" +"&BBS.Info(93)&" ，"&BBS.Info(121)&" +"&BBS.Info(94)&"，"&BBS.Info(122)&" +"&BBS.Info(95)&" 的奖励！"
			End If
			BBS.NetLog"主题管理：设置区置顶。<br>主题:"&left(Rs(1),20)&"<br>作者:"&SetUserName
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
		Content="<li>取消区置顶帖子---成功！</li>"
		If NotMe Then
			BBS.execute("update [user] set Coin=Coin-"&Int(BBS.Info(93))&",Mark=Mark-"&Int(BBS.Info(94))&",GameCoin=GameCoin-"&Int(BBS.Info(95))&" where name='"&SetUserName&"'")
			Content=Content&"<li>同时给该主题的作者："&SetUserName&" "&BBS.Info(120)&" -"&BBS.Info(93)&" ，"&BBS.Info(121)&" -"&BBS.Info(94)&"，"&BBS.Info(122)&" -"&BBS.Info(95)&" 的操作！</li>"
		End If
		BBS.NetLog"主题管理：取消区置顶。<br>主题:"&left(Rs(1),20)&"<br>作者:"&SetUserName
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
			Caption="错误信息"
			Content="该主题帖子已经是精华帖子了！"
		ELse
			BBS.Execute("update [Topic] set IsGood=1 where TopicID="&ID&" And boardid="&BBS.boardid&"")
			BBS.execute("update [User] set GoodNum=GoodNum+1 where name='"&SetUserName&"'")
			Content="<li>设定为精华帖子---成功！</li>"
		If NotMe Then
			BBS.execute("update [User] set Coin=Coin+"&Int(BBS.Info(99))&",Mark=Mark+"&Int(BBS.Info(100))&",GameCoin=GameCoin+"&Int(BBS.Info(101))&" where name='"&SetUserName&"'")
			Content=Content&"<li>同时给该主题的作者："&SetUserName&" "&BBS.Info(120)&" +"&BBS.Info(99)&" ，"&BBS.Info(121)&" +"&BBS.Info(100)&"，"&BBS.Info(122)&" +"&BBS.Info(101)&" 的奖励！</li>"
		End If
			BBS.NetLog"主题管理：设置精华。<br>主题:"&left(Rs(1),20)&"<br>作者:"&SetUserName
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
		S="单帖屏蔽"
	ELse
		Temp=0
		S="解除单帖屏蔽"
	End If
	BBS.execute("update [bbs"&BBS.TB&"] set IsDel="&Temp&" where BbsID="&BbsID&" And boardid="&BBS.boardid&"")
	Content="<li>"&S&"---成功！</li>"
	BBS.NetLog"管理帖子："&S
	Rs.close
End Sub

Sub SetNotTopicGood
	If SESSION(CacheName& "MyGradeInfo")(34)="0" Then BBS.GotoErr(70)
	Set Rs=BBS.Execute("Select IsGood,caption From[Topic] where TopicID="&ID)
	If Rs.eof Then
		BBS.GoToErr(58)
	Else
		IF Rs(0)=0 Then
			Caption="错误信息"
			Content="<li>该主题帖子已经被取消了精华了！</li>"
		ELse
			BBS.Execute("update [Topic] set IsGood=0 where TopicID="&ID)
			Content="<li>取消帖子精华---成功！</li>"
			If NotMe Then
				BBS.execute("update [User] set Coin=Coin-"&Int(BBS.Info(99))&",Mark=Mark-"&Int(BBS.Info(100))&",GameCoin=GameCoin-"&Int(BBS.Info(101))&",GoodNum=GoodNum-1 where name='"&SetUserName&"'")
				Content=Content&"<li>同时给该主题的作者："&SetUserName&" "&BBS.Info(120)&" -"&BBS.Info(99)&" ，"&BBS.Info(121)&" -"&BBS.Info(100)&"，"&BBS.Info(122)&" -"&BBS.Info(101)&" 的操作！</li>"
			End If
			BBS.NetLog"主题管理：取消精华。<br>主题:"&left(Rs(1),20)&"<br>作者:"&SetUserName
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
			Caption="错误信息"
			Content="<li>该主题帖子已经被锁定了！</li>"
		Else
			BBS.execute("update [Topic] set IsLock=1 where TopicID="&ID&" And boardid="&BBS.boardid&"")
			Content="<li>帖子锁定---成功！</li>"
			BBS.NetLog"主题管理：主题锁定。<br>主题:"&left(Rs(1),20)&"<br>作者:"&SetUserName
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
			Caption="错误信息"
			Content="<li>该主题帖子已经解锁了！</li>"
		Else
			BBS.execute("update [Topic] set IsLock=0 where TopicID="&ID&" And boardid="&BBS.boardid&"")
			Content="<li>帖子解锁---成功！</li>"
			BBS.NetLog"主题管理：主题解锁。<br>主题:"&left(Rs(1),20)&"<br>作者:"&SetUserName
		End IF
	End if
	Rs.Close
End Sub
Sub DelMy(IsTopic)
	Dim BbsID
	BbsID=BBS.CheckNum(request.querystring("BbsID"))
	'删除自己
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
				   TopicLastReply=RRs(0)&"|暂无回复"
				Else
				   TopicLastReply="|暂无回复"
				End If
		        RRs.CLose:Set RRs=Nothing
		End If
		ReRs.CLose:Set ReRs=Nothing
		BBS.execute("Update [Topic] set ReplyNum=ReplyNum-1,LastReply='"&TopicLastReply&"' where TopicId="&ID&"")
		UpdateSys 1,0
	End If
	Caption="删除成功"
	Content="已经成功删除了你自己发表的帖子！"
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
		If Rs(0)=ID Then'是主题
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
	'是主题
		If lcase(.MyName)=Lcase(SetUserName) Then
			Call DelMy(1)
			Exit Sub
		End If
	End if
	Cmd=Request("Cmd")
	If SESSION(CacheName& "MyGradeInfo")(26)="0" Then .GotoErr(70)
	If Cmd="del" or .Info(51)="0" then'快速删除
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
			Content="<li>请填写删除理由！<a href='javascript:history.go(-1)'>[返回]</a></li>"	
		ElseIf Len(Cause)>10 Then
			Content="<li>删除理由描述不能超过10个字符！<a href='javascript:history.go(-1)'>[返回]</a></li>"	
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
		                        TopicLastReply=RRs(0)&"|暂无回复"
		                    Else
		                        TopicLastReply="|暂无回复"
		                    End If
		                    RRs.CLose:Set RRs=Nothing
					End If
					ReRs.CLose:Set ReRs=Nothing
					.execute("Update [Topic] set ReplyNum=ReplyNum-1,LastReply='"&TopicLastReply&"' where TopicId="&ID&"")
					.execute("update [bbs"&.TB&"] set IsDel=1 where TopicID="&ID&" And ReplyTopicID=0 And BbsID="&BbsID&" And boardid="&.boardid&"")
					UpdateSys 1,0
				End If
				Temp=GetGained(Rs(1),Coin,Mark,GameCoin)
			'发信
				If IsSms="yes" Then
					Smss="你发表的帖子被删除："&Cause&vbcrlf&Temp
					If Sms<>"" Then Smss=Smss&vbcrlf&vbcrlf&"以下是操作人 "&.MyName&" 给你的附加留言信息："&vbcrlf&Sms
					.Execute("insert into [Sms](name,MyName,Content,MyFlag) values('自动送信系统','"&Rs(1)&"','"&Smss&"',1)")
					.Execute("update [User] set NewSmsNum=NewSmsNum+1,SmsSize=SmsSize+1 Where Name='"&Rs(1)&"'")
					.UpdageOnline Rs(1),1
				End If
				If Temp<>"" Then Temp="<li>"&Temp&"</li>"
				If BBSID=0 Then
					Content="<li>删除主题帖子---成功！</li>"&Temp
					.NetLog"删除主题帖:"&Cause&","&Temp&"<br>主题:"&left(Rs(2),20)&"<br>作者:"&SetUserName
				Else
					GotoUrl=True
					Content="<li>删除帖子---成功！</li>"&Temp
					.NetLog"删除回复帖:"&Cause&"<br>作者:"&SetUserName
				End If
				Rs.Close
			Else
				Caption="错误信息"
				Content="<li>帖子已经删除了！</li>"
			End IF
		End If
	Else
		IsShow=False
		If BBSID<>0 Then Caption="删除帖子" Else Caption="删除主题"
		S="<form method=POST  style=""margin:0px"" action='?action=删除&Cmd=del&TB="&.TB&"&id="&id&"&boardid="&.boardid&"&BBSID="&BBSID&"'>"
		S=S&Cue
		S=S&.Row("<b>操作理由：</b><select name='select' onChange='cause.value=this.options[this.selectedIndex].value'><option selected></option><option value='本版严禁广告'>本版严禁广告</option><option value='帖子内容违规'>帖子内容违规</option><option value='罚！无聊的乱灌水'>无聊的乱灌水</option><option value='重复发此类帖'>重复发此类帖</option></select>","<input name='cause' type='text' value='' size='30' maxlength='20'>必填，最多10个字符","65%","")
		S=S&.Row("<b>惩罚操作：</b>",.Info(120)&" <select name='coin'>"&Options(.Info(113),2)&"</select> "&.Info(121)&" <select name='mark'>"&Options(.Info(114),2)&"</select> "&.Info(122)&"<select name='gamecoin'>"&Options(.Info(115),2)&" </select>","65%","")
		S=S&.Row("<b>留言通知帖子作者：</b>","启用<input name='issms' onclick='if(sms.disabled==true){sms.disabled=false;sms.focus()}else{sms.disabled=true;}' type='checkbox' value='yes'>&nbsp; 留言附加信息：<input name='sms' size='30' type='text' value='' disabled='true'>","65%","")
		S=S&"<div style=""padding:2px;BACKGROUND: "&.SkinsPIC(2)&";"" align=""center""><input Class='button' type=""submit"" value=""确定操作"" /></div></form>"
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
	S=S&BBS.Row("请选择帖子要移动到的论坛：",GetBoardList(),"65%","")
	If Lcase(SetUserName)<>Lcase(BBS.MyName) Then
	S=S&BBS.Row("是否留言通知帖子作者：<input name='issms' onclick='if(sms.disabled==true){sms.disabled=false;sms.value=""通知：您的帖子被管理员("&BBS.MyName&")移动到这里：""}else{sms.disabled=true;sms.value="""";}' type='checkbox' value='yes'>","<input name='sms' size='50' class='text' type='text' value='' disabled='true'>","65%","")
	End If
	S=S&"<div style=""padding:2px;BACKGROUND: "&BBS.SkinsPIC(2)&";"" align=""center""><input Class='button' type=""submit"" value=""确定操作"" /></div></form>"
	BBS.ShowTable "移动帖子",S
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
		Sms=Sms&vbcrlf&"<a href=topic.asp?boardid="&Newboardid&"&id="&id&"&TB="&.TB&">请点击这里您的帖子</a>"
		.Execute("insert into [Sms](name,MyName,Content,MyFlag) values('自动送信系统','"&SetUserName&"','"&Sms&"',1)")
		.Execute("update [User] set NewSmsNum=NewSmsNum+1,SmsSize=SmsSize+1 Where Name='"&SetUserName&"'")
		.UpdageOnline SetUserName,1
	End If
	End If
	'更新版块
	Dim Boardupdate,LastReply
	Boardupdate=.GetEachBoardCache(.boardid)
	If Boardupdate(7)=ID&"" Then
	If .BoardString(6)=6 or .BoardString(6)=5 Then	'特殊版面
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
	Content="<li>移动帖子---成功！！</li>"
	.NetLog"主题管理：移动"
	End with
End Sub

Function GetBoardList()
	Dim Temp,i
	Temp="<select Style='font-size: 9pt' name='boardid' >"
	If IsArray(BBS.Board_Rs) Then
		For i=0 To Ubound(BBS.Board_Rs,2)
		IF BBS.Board_Rs(0,I)=1 Then
			Temp=Temp&"<option value="&BBS.Board_Rs(1,I)&">├"&BBS.Board_Rs(3,I)&"</option>"
		ElseIf BBS.Board_Rs(0,I)=2 Then
			Temp=Temp&"<option value="&BBS.Board_Rs(1,I)&">∣├"&BBS.Board_Rs(3,I)&"</option>"
		End If
		Next
	End If
	GetBoardList=Temp&"</select>"
End Function


Sub UpdateSys(EssayNum,TopicNum)
	with BBS
	Dim LastReply,TempContent,TempID,Rs1
	.execute("update [Config] set AllEssayNum=AllEssayNum-"&EssayNum&",TopicNum=TopicNum-"&TopicNum)	
	'如果是特殊版面不显示版块回复
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
	'更新父版块
	If .BoardDepth>1 Then
		.Execute("Update [Board] set LastReply='"&LastReply&"',EssayNum=EssayNum-"&EssayNum&",TopicNum=TopicNum-"&TopicNum&" where boardid In ("&.BoardParentStr&") And ParentID<>0")
		TempID=TempID&","&.BoardParentStr
	End If
	.UpdateEcachBoardCache TempID,-EssayNum&"|"&-TopicNum&"|0|"&LastReply
	'更新系统动态缓存数据
	TempContent=.InfoUpdate(0)-Int(EssayNum)&","&.InfoUpdate(1)-Int(TopicNum)&","&.InfoUpdate(2)&","&.InfoUpdate(3)&","&.InfoUpdate(4)&","&.InfoUpdate(5)&","&.InfoUpdate(6)&","&.InfoUpdate(7)&","&.InfoUpdate(8)&","&.InfoUpdate(9)&","&.InfoUpdate(10)
	.Cache.Add "InfoUpdate",TempContent,dateadd("n",2000,BBS.NowBBSTime)
	End with
End Sub

Sub TopHeight
	If SESSION(CacheName& "MyGradeInfo")(29)="0" Then BBS.GotoErr(70)
	BBS.Execute("update [Topic] set LastTime='"&BBS.NowBbsTime&"' where TopicID="&ID&" And boardid="&BBS.boardid&"")
	BBS.Execute("update [bbs"&BBS.TB&"] set LastTime='"&BBS.NowBbsTime&"' where TopicID="&ID&" And boardid="&BBS.boardid&"")
	Content="<Li>贴子主题提升---成功！！"
	BBS.NetLog"主题管理：提升"
End Sub

Sub Setsubside
	If SESSION(CacheName& "MyGradeInfo")(30)="0" Then BBS.GotoErr(70)
	BBS.Execute("update [Topic] set LastTime=LastTime-30 where TopicID="&ID&" And boardid="&BBS.boardid&"")
	Content="<Li>已经成功的使贴子主题沉底到一个月前新帖后面！"
	BBS.NetLog"主题管理：沉底"
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
	S="删除评帖记录 "
	If SESSION(CacheName& "MyGradeInfo")(41)="0" Then BBS.GotoErr(70)
	BbsID=BBS.CheckNum(request.querystring("BbsID"))
	BBS.Execute("delete from [appraise] where BbsID="&BbsID&" and TopicID="&ID)
	BBS.Execute("update [bbs"&BBS.TB&"] set IsAppraise=0 where BbsID="&BbsID&" And boardid="&BBS.boardid&"")
	Content="<li>"&S&"---成功！</li>"
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
		S=S&BBS.Row("<b>操作理由：</b><select name='select' onChange='cause.value=this.options[this.selectedIndex].value'><option selected>帖子评价自定义</option><option value='奖！此帖子不错哦'>奖！此帖子不错哦</option><option value='奖！感谢无私贡献'>奖！感谢无私贡献</option><option value='奖！好文章给奖励'>奖！好文章给奖励</option><option value='罚！本版严禁广告'>罚！本版严禁广告</option><option value='罚！帖子内容违规'>罚！帖子内容违规</option><option value='罚！无聊的乱灌水'>罚！无聊的乱灌水</option><option value='罚！重复发此类帖'>罚！重复发此类帖</option></select>","<input name='cause' type='text' value='' size='30' maxlength='25'>必填，最多22个字符","65%","")
		S=S&BBS.Row("<b>奖罚操作：</b>",BBS.Info(120)&" <select name='coin'>"&Options(BBS.Info(113),0)&"</select> "&BBS.Info(121)&" <select name='mark'>"&Options(BBS.Info(114),0)&"</select> "&BBS.Info(122)&"<select name='gamecoin'>"&Options(BBS.Info(115),0)&" </select>","65%","")
		S=S&BBS.Row("<b>留言通知帖子作者：</b>","启用<input name='issms' onclick='if(sms.disabled==true){sms.disabled=false;sms.focus()}else{sms.disabled=true;}' type='checkbox' value='yes'>&nbsp; 留言附加信息：<input name='sms' size='30' type='text' value='' disabled='true'>","65%","")
		S=S&"<div style=""padding:2px;BACKGROUND: "&BBS.SkinsPIC(2)&";"" align=""center""><input Class='button' type=""submit"" value=""确定操作"" /></div></form>"
	End If
	BBS.ShowTable"帖子评价",S
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
	Caption="评帖错误"
	Set Rs=BBS.execute("Select Name,Caption From [bbs"&BBS.TB&"] where BbsID="&BbsID&" And boardid="&BBS.boardid&"")
	IF Rs.eof Then
		BBS.GoToErr(58)
	ElseIf Lcase(Rs(0))=Lcase(BBS.MyName) Then
		Content="<li>不能对自己进行评帖！</li>"
	ElseIf Cause="" And (Mark=0 and Coin=0 and GameCoin=0) Then
		Content="<li>请填写完整再提交！</li>"	
	ElseIf Len(Cause)>22 Then
		Content="<li>评帖理由描述不能超过25个字符！</li>"	
	Else
		Cause=BBS.Fun.HtmlCode(Cause)
		BBS.execute("insert into [Appraise](BbsID,TopicID,Cause,Mark,Coin,GameCoin,AdminName,AddTime)VALUES("&BbsID&","&ID&",'"&Cause&"',"&Mark&","&Coin&","&GameCoin&",'"&BBS.MyName&"','"&BBS.NowBbsTime&"')")
		Temp=GetGained(Rs(0),Coin,Mark,GameCoin)
		If IsSms="yes" Then
			Smss="你的帖子：<a href="""&toUrl&""">"&Rs(1)&"</a><br>被评价："&Cause&"<br>"&Temp
			If Sms<>"" Then Smss=Smss&"<br><br>以下是操作人:"&BBS.MyName&" 给你的附加留言信息："&vbcrlf&Sms
			BBS.Execute("insert into [Sms](name,MyName,Content,MyFlag) values('自动送信系统','"&Rs(0)&"','"&Smss&"',1)")
			BBS.Execute("update [User] set NewSmsNum=NewSmsNum+1,SmsSize=SmsSize+1 Where Name='"&Rs(0)&"'")
			BBS.UpdageOnline Rs(0),1
		End If
		BBS.NetLog"帖子评价:"&Cause&","&Temp
		Rs.Close
		BBS.Execute("Update [bbs"&BBS.TB&"] Set IsAppraise=1 where BbsID="&BbsID&" And boardid="&BBS.boardid&"")
		Caption="帖子评定"
		If Temp<>"" Then Temp="<li>"&Temp&"</li>"
		Content="<li>帖子评定成功!</li>"&Temp
	End If
End Sub

Function GetGained(UserName,Coin,Mark,GameCoin)
	If Coin<>0 or Mark<>0 or GameCoin<>0 Then 
		GetGained="并且对作者 "&UserName&" 进行了"
		If Coin<>0 Then GetGained=GetGained&BBS.Info(120)&Coin&","
		If Mark<>0 Then GetGained=GetGained&BBS.Info(121)&Mark&","
		If GameCoin<>0 Then GetGained=GetGained&BBS.Info(122)&GameCoin
		GetGained=GetGained&"的操作。"
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
		Caption="投票帖子的标题："&BBS.execute("Select Caption from [Topic] where TopicID="&ID)(0)
		Vote=split(Rs(1),"|")
		VoteNum=Split(Rs(2),"|")
		II=UBound(Vote)
		For I = 1 To II
			S=S&BBS.Row("&nbsp;"&i,"<input size='80' name='Votes"&i&"' type='text' value='"&Vote(i)&"'> 投票数：<input size='3' name='VoteNum"&i&"' type='text' value='"&VoteNum(i)&"'>","95%","")
		Next
		S=S&BBS.Row("&nbsp;新建","<input size='80' name='Votes"&II+1&"' type='text' value=''> 投票数：<input size='3' name='VoteNum"&II+1&"' type='text' class='text' value='0'>","95%","")
		If Rs(3)=2 then VoteType=" checked"
		S="<form style=""margin:0px"" action='?action=SaveVote&TB="&BBS.TB&"&boardid="&BBS.boardid&"&id="&id&"' method='post'><input name='AutoValue' type='hidden' value='"&UBound(Vote)+1&"'>"&S&"<div style="" padding:5px;BACKGROUND: "&BBS.SkinsPIC(2)&";"" align=""center"">是否多选：<input name='VoteType' type='checkbox' value=2 "&voteType&"> &nbsp; 过期时间：<input size=20 name='OutTime' type='text' value='"&Rs(4)&"'> &nbsp; <input type='submit' value='确定修改'></div></form>"
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
		Temp="<li><a href='Topic.asp?TB="&BBS.TB&"&boardid="&BBS.boardid&"&id="&id&"'>回到投票帖子</a></li><li><a href='?action=EditVote&TB="&BBS.TB&"&boardid="&BBS.boardid&"&id="&id&"'>继续修改投票内容</a></li>"
	Else
		BBS.Execute("Delete From[TopicVote] where TopicID="&ID)
		BBS.Execute("Delete From[TopicVoteUser] where TopicID="&ID)
		BBS.Execute("Update [Topic] Set IsVote=False where TopicID="&ID)
		Temp="<li>已经成功删除的投票内容！<li><a href='topic.asp?TB="&BBS.TB&"&boardid="&BBS.boardid&"&id="&id&"'>回到主题帖子</a></li>"
	End IF
	BBS.NetLog"管理主题:修改投票选项"
	Caption="操作成功"
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
	BBS.execute("update [Topic] set Caption='【已解决】'&Caption where TopicID="&ID&" and boardid="&BBS.boardid&" and Caption not like'%【已解决】%'")
	Content="<li>设定帖子为已解决帖子---成功！！</li>"
	BBS.NetLog"设定帖子为已解决"
End Sub

Sub Affirm()
If Request.ServerVariables("request_method") <> "POST" then
Response.write "<form name='KK' method=post action=?"&Request.ServerVariables("Query_String")&"></form><SCRIPT>if(confirm('您确定要执行该操作?')){returnValue=KK.submit()}else{returnValue=history.back()}</SCRIPT>"
Response.End
End If
End Sub
%>