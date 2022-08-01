<!--#include file="Inc.asp"-->
<%Dim Action,ID
BBS.CheckMake
If not BBS.FoundUser Then BBS.GoToErr(10)
ID=BBS.CheckNum(request.querystring("Id"))
Action=Lcase(Request.querystring("Action"))
If len(Action)>13 then BBS.GoToErr(1)
If Action="vote" Then
	SubmitVote()
Else
	SubmitBuyer()
End If
Set BBS =Nothing
Response.redirect(Request.ServerVariables("HTTP_REFERER"))

Sub SubmitVote()
	Dim Temp,Rs,i,VoteType,Vote,VoteNum,TempVote,MyOpt,OutTime,AllVoteNum
	IF ID=0 Then BBS.GoToErr(1)
	VoteType=BBS.checknum(request.querystring("type"))
	Set Rs=BBS.execute("select Vote,VoteNum,VoteType,OutTime From[TopicVote] where TopicID="&ID&"")
	IF Not Rs.Eof Then
		Vote=split(Rs("vote"),"|")
		VoteNum=split(Rs("voteNum"),"|")
		OutTime=Rs("OutTime")
		IF VoteType<>Int(Rs("VoteType")) Then BBS.GoToErr(1)
		TempVote=Vote
		if VoteType=1 then
			Temp=BBS.CheckNum(Request.form("opt"))
			MyOpt=Temp
			For i=1 to ubound(Vote)
				If i=Temp then VoteNum(i)=VoteNum(i)+1
				AllVoteNum=AllVoteNum&"|"&VoteNum(i)
			Next
		ElseIf VoteType=2 Then
			Temp=0
			TempVote=Vote
			For i=1 to ubound(Vote)
				TempVote(i)=BBS.Checknum(Request.form("opt"&i&""))
				Temp=TempVote(i)+Temp
				IF TempVote(i)=0 Then TempVote(i)=VoteNum(i)
				IF TempVote(i)=i Then
				 TempVote(i)=Votenum(i)+1
				 MyOpt=MyOpt&","&i
				End if
				AllVoteNum=AllVoteNum&"|"&TempVote(i)  
			Next
		Else
			BBS.GoToErr(1)
		End if
		If Temp=0 Then BBS.alert"您还没有选择投票项目！","back"
		IF Temp<>0 And BBS.execute("select User From [TopicVoteUser] where User='"&BBS.MyName&"' and TopicID="&ID&"").Eof Then
			If DateDiff("s",BBS.NowBbsTime,OutTime)>0 then
			BBS.execute("update [TopicVote] Set VoteNum='"&AllvoteNum&"' where TopicID="&ID&"")
			BBS.execute("update [Topic] Set LastTime='"&BBS.NowBbsTime&"' where TopicID="&ID&"")
			BBS.execute("update [bbs"&BBS.TB&"] Set LastTime='"&BBS.NowBbsTime&"' where TopicID="&ID&"")
			BBS.execute("Insert into [TopicVoteUser](TopicID,[User],VoteNum)VALUES("&ID&",'"&BBS.MyName&"','"&MyOpt&"')")
			End If
		End If
	End if
	Rs.Close
	Set Rs=nothing
End Sub
Sub SubmitBuyer()
	If ID=0 Then BBS.GoToErr(1)
	Dim Temp,Rs,Rss,Buyer,re,str
	Set Rs=BBS.execute("Select Content,Name From[bbs"&BBS.TB&"] where BbsID="&ID&"")
	IF Rs.eof Then BBS.GoToErr(1)
	Rss=Rs.GetRows(1)
	Rs.Close
	Temp=Replace(Rss(0,0),chr(10),"")
	Temp=Replace(Temp,chr(10),"")
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	re.Pattern="(^.*)(\[BUYPOST=*([0-9]*)\])(.*)(\[\/BUYPOST\])(.*)"
	Str=re.Replace(Temp,"$3")
	Set re=Nothing
	If isnumeric(Str) Then
		Str=int(Str)
	Else
		Str=0
	End if
	If Lcase(Rss(1,0))=Lcase(BBS.MyName) Then BBS.alert"您不能向自己购买！","back"
	If Int(SESSION(CACHENAME & "MyInfo")(7))<Str Then BBS.alert"钱不够，请再努力发帖赚钱吧！","back"
	Set Rs=BBS.Execute("select BBSID,UserName From [Buyer] where BBSID="&ID)
	If Not Rs.Eof Then
		If instr("|"&Lcase(Rs(1))&"|","|"&Lcase(BBS.MyName)&"|")>0 Then BBS.alert"您已经购买过了呀？","back"
		Temp=Rs(1)&"|"&BBS.MyName
		BBS.execute("Update [Buyer] Set UserName='"&Trim(Temp)&"' Where BbsID="&ID)
	Else
		BBS.Execute("insert into[Buyer](BBSID,UserName)values("&ID&",'"&BBS.MyName&"')")
	End IF
	BBS.execute("update [user] set Coin=Coin-"&Str&" where name='"&BBS.MyName&"'")
	BBS.execute("update [user] set Coin=Coin+"&Str&" where name='"&Rss(1,0)&"'")
	Session(CacheName & "MyInfo") = Empty
End Sub

%>