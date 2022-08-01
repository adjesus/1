<!--#include file="Inc.asp"-->
<%
Dim Action,Caption,Content,Face,Page,TmpUbbString
If Not BBS.Founduser then BBS.GoToerr(10)
BBS.CheckMake
IF (Session(CacheName&"SayTime")+Int(BBS.Info(11))/86400)>Now() Then BBS.Alert"本论坛为了防止灌水，限制了同一人发帖的时间间隔为"& BBS.Info(11)&" 秒！","back"
IF (Session("iCode")<>request.form("iCode") or Session("iCode")="") And BBS.Info(15)="1" then BBS.GoToErr(35)
BBS.CheckBoard()
Face=BBS.CheckNum(Request.form("face"))
Caption=Trim(BBS.Fun.filtrateHtmlCode(BBS.Fun.Checkbad(Request.form("Caption"))))
Content=BBS.Fun.Checkbad(BBS.Fun.GetForm("content"))

If Caption="" Or Content="" then BBS.GoToErr(27)

If BBS.Fun.CheckIsEmpty(Content) Then BBS.GoToErr(50)
If BBS.Info(60)="1" Then Content=BBS.Fun.Replacehtml(Content)

IF Len(Caption)>200 Then BBS.GoToErr(28)
IF Len(Content)>int(Session(CacheName & "MyGradeInfo")(9)) Then BBS.GoToErr(29)
TmpUbbString=BBS.Fun.UbbString(content)
BBS.Head"","",""

Action=lcase(request.querystring("action"))
If Len(Action)>10 Then BBS.GotoErr(1)
IF Action="reply" Then
 	Reply()
ElseIf Action="edit" Then
	Edit()
Else
	Say()
End if

If BBS.Info(15)=0 then Session("iCode")=Empty
Content="<div style=""margin:18px;line-height:150%"">"&Content&"</div>"
BBS.ShowTable Caption,Content
Session(CacheName & "SayTime")=Now()
BBS.Footer()
Set BBS =Nothing

Function CheckUploadType(Str)
	dim re,s
	s=Str
	Set re=new RegExp
	re.IgnoreCase=true
	re.Global=True
	re.Pattern="(^.*)\[upload=(.+?),(.+?)\](.+?)\[\/upload\](.*)"
	s=Re.replace(s,"$2")
	Set re=nothing
	CheckUploadType=s
End function


Sub Say()
	with BBS
	Dim ISvote,VoteType,VoteAutoValue,Votes,VoteNum,Outtime
	Dim UploadType,TopicLastReply,BoardLastReply,MaxID,TopicOpt
	Dim Temp,I,Font_S,Font_C
	IsVote=0
	Temp=CheckUploadType(Content)
	If Temp <> "" Then
		If instr(lcase("|"&.Info(34)&"|"&.Info(35)&"|"),lcase("|"&Temp&"|"))>0 then
			UploadType=Temp
		Else
			UploadType=""
		End if
	End IF
	VoteType=.CheckNum(request.Form("votetype"))
	If VoteType<>2 Then VoteType=1
	VoteAutoValue=.CheckNum(request.Form("autovalue"))
	For i=1 to VoteAutoValue
	Temp=Left(.Fun.Checkbad(Trim(.Fun.GetStr("Votes"&I))),250)
	IF Temp<>"" Then
		Votes=Votes&"|"&Temp
		VoteNum=VoteNum&"|0"
	End If
	Next
	Outtime=cDate(.NowBbsTime)+int(request.form("OutTime"))
	If Votes<>"" Then IsVote=1
	'标题样式
	Font_S=.FUN.GetStr("font_s")
	Font_C=.FUN.GetStr("font_c")
	If Font_S<>"" or Font_C<>"" Then Temp=Font_S&"|"&Font_C Else Temp=""
	'插入主题表
	TopicLastReply=.Myname&"|"&.Fun.StrLeft(.Fun.FixReply(Content),40)
	.Execute("Insert into [Topic](Caption,Name,Face,IsVote,AddTime,LastTime,Lastreply,UploadType,boardid,SqlTableID,Font)Values('"&Caption&"','"&.Myname&"',"&Face&","&IsVote&",'"&.NowBbsTime&"','"&.NowBbsTime&"','"&TopicLastReply&"','"&UploadType&"',"&.boardid&","&.TB&",'"&Temp&"')")
	'得到新发表主题的ID	
	MaxID=.Execute("select Max(TopicID) from [Topic]")(0)
	'操作选项
	Call Topict(MaxID)
	'插入投票
	IF IsVote=1 Then .Execute("insert into[TopicVote](TopicID,Vote,VoteNum,VoteType,OutTime)values("&MaxID&",'"&votes&"','"&VoteNum&"',"&votetype&",'"&Outtime&"')")
	'插入帖子
	.Execute("Insert into [bbs"&.TB&"](TopicID,Caption,name,Content,AddTime,LastTime,Face,boardid,UbbString,IP)VALUES("&MaxID&",'"&Caption&"','"&.MyName&"','"&Content&"','"&.NowBbsTime&"','"&.NowBbsTime&"',"&Face&","&.boardid&",'"&TmpUbbString&"','"&.MyIP&"')")
	'更新版块
	If .BoardString(6)=1 or .BoardString(5)=1 Then'特殊版面
		BoardLastReply=""
	Else
		BoardLastReply=.MyName&"|"&.Fun.StrLeft(Caption,22)&"|"&.NowBbsTime&"|"&Face&"|"&MaxID&"|"&.boardid&"|"&.TB
	End If
		.Execute("Update [Board] set TopicNum=TopicNum+1,TodayNum=TodayNum+1,EssayNum=EssayNum+1,LastReply='"&BoardLastReply&"' where boardid="&.boardid&" And Depth>0")
		Temp=.boardid
		'更新上级版块
		If .BoardDepth>1 Then
			.Execute("Update [Board] set TopicNum=TopicNum+1,TodayNum=TodayNum+1,EssayNum=EssayNum+1,LastReply='"&BoardLastReply&"' where boardid In ("&.BoardParentStr&") And Depth>0")
			Temp=Temp&","&.BoardParentStr
		End If
		'更新各版块缓存
		.UpdateEcachBoardCache Temp,"1|1|1|"&BoardLastReply
	'更新系统
	.Execute("Update [config] set Topicnum=Topicnum+1,allEssayNum=AllEssayNum+1,TodayNum=TodayNum+1")	
	'更新用户
	If Len(Content)>Int(.info(112)) Then
		Temp="Coin=Coin+"&.Info(102)&",Mark=Mark+"&.Info(103)&",GameCoin=GameCoin+"&.Info(104)&","
		If .Info(76)="1" Then Content=ShowGained(.Info(102),.Info(103),.Info(104))
	Else
		Temp=""
		Content=""
	End If
		.execute("Update [user] set "&Temp&"EssayNum=EssayNum+1 WHERE ID="&.MyID&"")
	'更新等级
    if int(Session(CacheName & "MyGradeInfo")(1))=0 then
	.UpdateGrade .MyID,Session(CacheName & "MyInfo")(4)+1,int(Session(CacheName & "MyGradeInfo")(1))
	End If
	UpdateInfoCache(1)
	Caption="发帖成功！"
	Content="<meta http-equiv=""refresh"" content=""5;url=topic.asp?id="&MaxId&"&boardid="&.boardid&"&TB="&.TB&""" /><li><b>发表成功----如果你不点击下面的连接，将在 5 秒后自动跳转到 回到你所发的帖的页面！</b></li><li><a href=""topic.asp?id="&MaxId&"&boardid="&.boardid&"&TB="&.TB&""">回到你所发的帖的页面！</a></li><li><a href=""board.asp?boardid="&.boardid&""">"&.Boardname&"</a><li><a href=""index.asp"">"&.Info(0)&" 首页</a></li>"&Content
	End with
End Sub

Sub Reply()
	With BBS
	Dim Rs,ID,TopicUserName,TopicLastReply,BoardLastReply,Temp,TopicCoin
	ID=.Checknum(request.querystring("ID"))
	If Id=0 or .boardid=0 Then .GoToErr(1)
	'读主题信息
	Set Rs=.Execute("Select Name,IsLock,SqlTableID,boardid,Caption From [Topic] where TopicID="&ID&" And (boardid="&.boardid&" Or TopType=4 or TopType=5) And SqlTableID="&.TB&" And IsDel=0 ")
	IF Rs.Eof Then .GoToErr(21)  
	IF Rs(1)=1 Then .GoToErr(22)
	TopicUserName=Rs(0)
	.TB=Rs(2)
	.boardid=Rs(3)
	Caption="Re："&Rs(4)
	Rs.Close
	Set Rs=Nothing
	'插入帖子
	.execute("insert into [bbs"&.TB&"](ReplyTopicID,name,Caption,Content,AddTime,LastTime,Face,boardid,UbbString,ip)VALUES("&ID&",'"&.MyName&"','"&Caption&"','"&Content&"','"&.NowBbsTime&"','"&.NowBbsTime&"',"&face&","&.boardid&",'"&TmpUbbString&"','"&.MyIp&"')")
	'更新主题帖子
	.execute("Update [bbs"&.TB&"] Set LastTime='"&.NowBbsTime&"' where TopicId="&ID&"")
	'更新主题
	TopicLastReply=.MyName&"|"&.Fun.StrLeft(.Fun.FixReply(Content),40)
		.execute("Update [Topic] set ReplyNum=ReplyNum+1,LastTime='"&.NowBbsTime&"',LastReply='"&TopicLastReply&"' where TopicId="&ID&"")
	'更新楼主
	If Lcase(.MyName)<>Lcase(TopicUserName) Then
		.execute("Update [user] set Coin=Coin+"&.Info(111)&" WHERE Name='"&TopicUserName&"'")
	End If
	'更新版块
	If .BoardString(6)=6 or .BoardString(6)=5 Then	'特殊版面
		BoardLastReply=""
	Else
		If .Info(81)="1" Then
			Temp=.Fun.StrLeft(.Fun.FixReply(Content),22)
		Else
			Temp=.Fun.StrLeft(Caption,22)	
		End If	
		BoardLastReply=.MyName&"|"&Temp&"|"&.NowBbsTime&"|"&Face&"|"&ID&"|"&.boardid&"|"&.TB
	End If
		.execute("Update [Board] set lastReply='"&BoardLastReply&"',TodayNum=TodayNum+1,EssayNum=EssayNum+1 where boardid="&.boardid&" and Depth<>0")
		Temp=.boardid
		If .BoardDepth>1 Then
			.Execute("Update [Board] set TopicNum=TopicNum+1,TodayNum=TodayNum+1,EssayNum=EssayNum+1,LastReply='"&BoardLastReply&"' where boardid In ("&.BoardParentStr&") And Depth>0")
			Temp=Temp&","&.BoardParentStr
		End If
	.UpdateEcachBoardCache Temp,"1|0|1|"&BoardLastReply

	'更新系统
	.execute("Update [Config] set TodayNum=TodayNum+1,AllEssayNum=AllEssayNum+1")
	'更新用户
	If Len(Content)>Int(.info(112)) Then
		Temp="Coin=Coin+"&.Info(105)&",Mark=Mark+"&.Info(106)&",GameCoin=GameCoin+"&.Info(107)&","
		If .Info(76)="1" Then Content=ShowGained(.Info(105),.Info(106),.Info(106))
	Else
		Temp=""
		Content=""
	End If
	.execute("Update [user] set "&Temp&"EssayNum=EssayNum+1 WHERE ID="&.MyID)
	if int(Session(CacheName & "MyGradeInfo")(1))=0 then
	.UpdateGrade .MyID,Session(CacheName & "MyInfo")(4)+1,int(Session(CacheName & "MyGradeInfo")(1))
	End If
	UpdateInfoCache(0)
	'清分页缓存
	Response.Cookies(CacheName&"P")("Show_"&ID)=""
	Caption="回复成功 ！"
	Content="<meta http-equiv=refresh content=""5;url=topic.asp?id="&ID&"&boardid="&.boardid&"&TB="&.TB&"&page=999"" /><li><b>发表成功----如果你不点击下面的连接，将在 5 秒后自动跳转到 您所回复的帖子内容页面！</b></li><li><a href='topic.asp?id="&ID&"&boardid="&.boardid&"&TB="&.TB&"'>回到您所回复主题！</a></li><li><a href='topic.asp?id="&ID&"&boardid="&.boardid&"&TB="&.TB&"&page=999'>回到您所回复的帖子内容页面！</a></li><li><a href='board.asp?boardid="&.boardid&"'>"&.Boardname&"</a><li><a href='index.asp'>"&.Info(0)&" 首页</a>"&Content
	End with
End Sub

Sub Edit()
	Dim Temp,Rs,ID,BbsID,TopicID,EditChalk,ReplyTopicID,Font_S,Font_C
	With BBS
	Page=.CheckNum(request.querystring("page"))
	EditChalk=Request.form("editchalk")
	ID=.CheckNum(request.querystring("ID"))
	BbsID=.CheckNum(request.querystring("BbsID"))
	If BbsID=0 Or ID=0 Then .GoToErr(1)
	If EditChalk<>"No" Then
		Content=Content&vbcrlf&vbcrlf&"<div style=""color:#999999;text-align:right"">「该帖子被 "&.MyName&" 在 "&.NowBbsTime&" 编辑过」</div>"
	End If
	Set Rs=.Execute("select TopicID,ReplyTopicID,Caption,Name from [bbs"&.TB&"] where BbsID="&BbsID&" and IsDel=0")
	If Not Rs.eof  Then
		If Session(CacheName & "MyGradeInfo")(24)="0" And Lcase(.MyName)<>Lcase(rs("name")) Then .GoToErr(33)
		TopicID=Rs(0)
		ReplyTopicID=Rs(1)
		Temp=Rs(2)
	Else
		.GoToErr(58)
	End if
	Rs.Close
	If ReplyTopicID=0 then
		'标题样式
		Font_S=.FUN.GetStr("font_s")
		Font_C=.FUN.GetStr("font_c")
		If Font_S="" And Font_C="" Then
			Font_S=""
		Else
			If Font_S="no" Then Font_S=""
			If Font_C="no" Then Font_C=""
			If Font_S<>"" or Font_C<>"" Then
				Font_S=",Font='"&Font_S&"|"&Font_C&"'"
			Else
				Font_S=",Font=''"
			End If
		End If
		.execute("Update [Topic] set Caption='"&Caption&"',Face="&Face&",LastTime='"&.NowBbsTime&"'"&Font_S&" where TopicID="&TopicID&"")
	Else
		Caption=Temp
	End If
	'更新
	.execute("Update [bbs"&.TB&"] set Caption='"&Caption&"',Content='"&Content&"',Face="&Face&",LastTime='"&.NowBbsTime&"',UbbString='"&TmpUbbString&"',IP='"&.MyIp&"' where BbsID="&BbsID&"")
	'更新主题的最后回复
	dim tBBSID,tName,tLastTime,tFace
	Temp=""
	Set Rs=.execute("select top 1 BbsID,Name,Caption,Content,LastTime,Face from [bbs"&.TB&"] where boardid="&.boardid&" And (TopicID="&ID&" or ReplyTopicID="&ID&") And IsDel=0 order by BbsID desc")
	If Not Rs.Eof Then
		If BbsID=Rs(0) Then	Temp=Rs(1)&"|"&.Fun.StrLeft(.Fun.FixReply(Rs(2)),40)
		tBBSID=Rs(0)
		tName=Rs(1)
		Caption=Rs(2)
		Content=Rs(3)
		tLastTime=Rs(4)
		tFace=Rs(5)
	Else
		Temp="|"&.Fun.StrLeft(.Fun.FixReply(Content),40)
	End If
		Rs.Close
	Set Rs=Nothing
	If Temp<>"" then .execute("Update [Topic] set LastReply='"&Temp&"' where TopicId="&ID&"")
	'更新版块
	Dim Boardupdate,BoardLastReply
	Boardupdate=.GetEachBoardCache(.boardid)
	If Boardupdate(7)=ID&"" Then
		If .BoardString(6)=6 or .BoardString(6)=5 Then	'特殊版面
		Else
			If .Info(81)="1" Then
				Temp=.Fun.StrLeft(.Fun.FixReply(Content),22)
			Else
				Temp=.Fun.StrLeft(Caption,22)	
			End If	
			BoardLastReply=tName&"|"&Temp&"|"&tLastTime&"|"&tFace&"|"&ID&"|"&.boardid&"|"&.TB
			.execute("Update [Board] set lastReply='"&BoardLastReply&"' where boardid="&.boardid&" and Depth<>0")
			Temp=.boardid
			If .BoardDepth>1 Then
				.Execute("Update [Board] set LastReply='"&BoardLastReply&"' where boardid In ("&.BoardParentStr&") And Depth>0")
				Temp=Temp&","&.BoardParentStr
			End If
			.UpdateEcachBoardCache Temp,"0|0|0|"&BoardLastReply
		End If
	End If
	Caption="编辑帖子"
	Content="<li>修改成功！ <a href='topic.asp?ID="&ID&"&boardid="&.boardid&"&TB="&.TB&"&page="&page&"'>回到帖子页面</a></li><li><a href='board.asp?boardid="&.boardid&"'>"&.BoardName&"</a></li><li> <a href='index.asp'>"&.Info(0)&"论坛首页</a></li>"
	End with
End Sub

'更新缓存(参数：0回复/1主题)
Sub UpdateInfoCache(IsSay)
	Dim Temp,Max
	Temp=BBS.Infoupdate(2)+1
	Max=BBS.InfoUpdate(4)
	If Int(Temp)>Int(Max) Then
		BBS.Execute("Update [Config] set MaxEssayNum="&Temp&"")
		Max=Temp
	End If
	Temp=Replace(Join(BBS.InfoUpdate,","),BBS.InfoUpdate(0)&","&BBS.InfoUpdate(1)&","&BBS.InfoUpdate(2)&","&BBS.InfoUpdate(3)&","&BBS.InfoUpdate(4)  ,	BBS.InfoUpdate(0)+1&","&BBS.InfoUpdate(1)+Int(IsSay)&","&BBS.InfoUpdate(2)+1&","&BBS.InfoUpdate(3)&","&Max)
	BBS.Cache.Add "InfoUpdate",Temp,dateadd("n",2000,BBS.NowBBSTime)
	Session(CacheName & "MyInfo") = Empty
End Sub
'显示发帖奖励
Function ShowGained(C,M,G)
	If C<>"0" Then ShowGained=BBS.Info(120)&":<span style='color:#F00'>"&C&"</span> "
	If M<>"0" Then ShowGained=ShowGained&BBS.Info(121)&":<span style='color:#F00'>"&M&"</span> "
	If G<>"0" Then ShowGained=ShowGained&BBS.Info(122)&":<span style='color:#F00'>"&G&"</span> "
	ShowGained="<li>您本次发帖获得："&ShowGained&"</li>"
End Function
'主题选项
Sub Topict(M_ID)
	Dim S
	S=""
	If request.Form("top")="1" Then S="toptype=3,"
	If Request.Form("classtop")="1" Then S="toptype=4,"
	If Request.Form("alltop")="1" Then S="toptype=5,"
	If Request.Form("good")="1" Then S=S&"isgood=1,"
	If Request.Form("lock")="1" Then S=S&"islock=1,"
	If S<>"" Then
	S=Left(S,len(S)-1)
	BBS.Execute("update [Topic] Set "&S&" where TopicID="&M_ID)
	End If
End Sub
%>
