<!--#include file="Inc.asp"-->
<!--#include file="Inc/Page_Cls.asp"-->
<%
Dim PageString,strPageInfo,Page_Url,Temp
If Request("find")<>"" then
	Response.Redirect"search.asp?key="&Request("find")&"&SType=2&STime=0&tb="&BBS.tb&"&boardid="&BBS.boardid&"&Action=topic"
	Response.End
End IF
BBS.CheckBoard()
If Request.QueryString("page") > 1 Then
  Page_Url = "&Page="&Request.QueryString("page")
Else
  Page_Url = ""
End If
BBS.Head "board.asp?boardid="&BBS.boardid&Page_Url,BBS.Boardname,""
If BBS.BoardChild>0 Then ShowBoard()
If BBS.BoardDepth>0 And BBS.BoardString(0)="0" Then
  PageString=ShowTopicList
  Response.Write(PageString)
Else

End If
	Temp="<iframe name='hiddenframe' frameborder='0' height='0' id='hiddenframe'></iframe>"
	If Session(CacheName&"online")="1" Then 
		Temp=Replace(Temp,"id="," src='online.asp?boardid="&BBS.boardid&"&id=1' id=")
	End If
	response.write Temp
BBS.Footer()
Set BBS =Nothing

Sub ShowBoard()
	Dim Board,Rs,i,BigBoard,BoardChild,BoardStr,II
	Dim area,area2,Child,width
	With BBS
	If Not IsArray(.Board_Rs) Then .GetboardCache()
	If Not IsArray(.Board_Rs) Then Exit Sub
	area=.ReadSkins("版块分区表格")
	area2=.ReadSkins("版块分区简洁表格")
	For i=0 To Ubound(BBS.Board_Rs,2)
		If .Board_Rs(1,i)=.boardid Then
			If .BoardString(1)="1" or Session(CacheName& "BoardStyle")="1" Then
				BigBoard=area2
				Child=Int(.Board_Rs(8,i))
				If Child<Int(.BoardString(2)) Then
					Width=100\Child
				Else
					width=100\Int(.BoardString(2))
				End If
				II=0
			Else
				BigBoard=area
			End If
				BoardStr=Replace(BigBoard,"{分类名称}",.Board_Rs(3,i))
				BoardStr=Replace(BoardStr,"{分类ID}",.Board_Rs(1,i))
		ElseIf .Board_Rs(2,i)=.boardid Then
			If .BoardString(1)="1" or Session(CacheName& "BoardStyle")="1" Then
			    Child=Child-Int(.Board_Rs(8,i))
				II=II+1
				If II=Int(.BoardString(2)) or II=Child Then
					BoardChild=BoardChild&"<div style='float:left;max-width:"&width&"%'>"&.GetboardInfo("1",i)&"</div>"
				Else
					BoardChild=BoardChild&"<div style='float:left;width:"&width&"%'>"&.GetboardInfo("1",i)&"</div>"
				End If
			Else
				BoardChild=BoardChild&.GetboardInfo("0",i)
			End IF
		End If
	Next
	BoardStr=Replace(BoardStr,"{显示版块}",BoardChild)
	Board = Board&BoardStr
	Response.Write Board
	End With
End Sub

Function ShowTopicList()
	Dim S,Button,S1
	Button=""
	IF BBS.BoardString(4)="0" or  BBS.MyAdmin=9 Or (BBS.MyAdmin=7 And BBS.IsBoardAdmin) Then
	Button="<a href='post.asp?boardid="&BBS.boardid&"'>"&BBS.SkinsPIC(7)&"</a> <a href='post.asp?action=vote&boardid="&BBS.boardid&"'>"&BBS.SkinsPIC(8)&"</a>"
	End If
	S=BBS.ReadSkins("主题列表表格")
	S=Replace(S,"{发帖按钮}",Button)
	S=Replace(S,"{版主}",BBS.Boardadmin)
	S=Replace(S,"{版块ID}",BBS.boardid)
	S=Replace(S,"{数据表ID}",BBS.tb)
	S=Replace(S,"{公告}",BBS.Placard(BBS.boardid))
	S=Replace(S,"{版块名称}",BBS.Boardname)
	If BBS.Info(21)="1" then S=Replace(S,"{显示在线}",ShowOnline()) Else S=Replace(S,"{显示在线}","")
	S=Replace(S,"{版块下拉列表}","<script language=""JavaScript"" type=""text/javascript"">BoardSelect()</script>")
	S1=ShowTopic()
	S=Replace(S,"{分页}",strPageInfo)
	S=Replace(S,"{显示主题列表}",S1)
	ShowTopicList=S
End Function

Function Showonline()
	Dim S,Boardupdate
	S=BBS.ReadSkins("显示版块在线")
	S=Replace(S,"{显示在线列表}","<div id=""showon""></div>")
	If Session(CacheName&"Online")="1" Then S=Replace(S,"{展开在线列表图标}","-.gif") Else S=Replace(S,"{展开在线列表图标}","+.gif")	
	S=Replace(S,"{版块ID}",BBS.boardid)
	S=Replace(S,"{在线总数}",BBS.AllOnlineNum)
	S=Replace(S,"{本版在线总数}",BBS.BoardOnlineNum)
	S=Replace(S,"{本版在线会员数}",BBS.BoardUserOnlineNum)
	S=Replace(S,"{本版在线游客数}",Int(BBS.BoardOnlineNum)-int(BBS.BoardUserOnlineNum))
	S=Replace(S,"{管理员}",BBS.SkinsPic(21))
	S=Replace(S,"{超级版主}",BBS.SkinsPic(22))
	S=Replace(S,"{版主}",BBS.SkinsPic(23))
	S=Replace(S,"{VIP会员}",BBS.SkinsPic(24))
	S=Replace(S,"{会员}",BBS.SkinsPic(25))
	S=Replace(S,"{隐身}",BBS.SkinsPic(26))
	S=Replace(S,"{游客}",BBS.SkinsPic(27))
	Boardupdate=BBS.GetEachBoardCache(BBS.boardid)
	S=Replace(S,"{今日帖数}",Boardupdate(2))
	ShowOnline=s
End Function

Function ShowTopic()
	Dim S,intPageNow,arr_Rs,i,P,Conut,page,Flag,Condection,TopicLine
	Dim TopicS,Caption,Facepic,Moodpic,LastRe,RePageUrl,UploadType,RePage,leftn,ii
	intPageNow = Request.QueryString("page")
	Condection= "(boardid="&BBS.boardid&" or TopType=5 or (TopType=4 and boardid in ("&BBS.BoardRoots&"))) And IsDel=0"
	Flag=BBS.CheckNum(Request.QueryString("Flag"))
	If Flag=1 Then Condection=Condection&" And IsGood=1"
	If Flag=2 Then Condection=Condection&" And DATEDIFF('d',[LastTime],'"&BBS.NowBbsTime&"')<1"
	If Flag=3 Then Condection=Condection&" And Name='"&BBS.MyName&"'"
	Set P = New Cls_PageView
	P.strTableName = "[Topic]"
	P.strPageUrl = "?flag="&flag&"&boardid="&BBS.boardid
	P.strFieldsList = "TopicID,Face,Caption,Name,TopType,IsGood,AddTime,boardid,LastTime,Hits,LastReply,UploadType,IsVote,ReplyNum,SqlTableID,IsLock,Font"
	P.strCondiction = Condection
	P.strOrderList = "TopType desc,LastTime desc"
	P.strPrimaryKey = "TopicID"
	P.intPageSize = Int(BBS.Info(61))
	P.intPageNow = intPageNow
	P.strCookiesName = "List"&BBS.boardid&"BBS"&Flag'客户端记录总数
	'P.Reloadtime=3'默认三分钟更新Cookies
	P.InitClass
	Arr_Rs = P.arrRecordInfo
	strPageInfo = P.strPageInfo
	Set P = nothing
	If IsArray(Arr_Rs) Then
	TopicLine = -1
	For i = 0 to UBound(Arr_Rs, 2)
		Moodpic=BBS.SkinsPIC(16)
		If Arr_Rs(13,i) > Int(BBS.Info(62)) Then Moodpic=BBS.SkinsPIC(15)
		If Arr_Rs(5,i)=1 Then Moodpic=BBS.SkinsPIC(13)'精华
		If Arr_Rs(15,i)=1 Then Moodpic=BBS.SkinsPIC(17)'锁定
		If Arr_Rs(12,i)=1 Then Moodpic=BBS.SkinsPIC(14)'投票
		If Arr_Rs(4,i)=5 Then Moodpic=BBS.SkinsPIC(10)'总顶
		If Arr_Rs(4,i)=4 Then Moodpic=BBS.SkinsPIC(11)'区顶
		If Arr_Rs(4,i)=3 Then Moodpic=BBS.SkinsPIC(12)'顶
		Facepic="<img src='pic/face/"&Arr_Rs(1,i)&".gIf' alt='' />"
		UploadType=""
		If Arr_Rs(11,i)<>"" Then Uploadtype="<img src='pic/FileType/"&Arr_Rs(11,i)&".gif' border='0' atl='"&Arr_Rs(11,i)&"' /> "
		LastRe=split(Arr_Rs(10,i),"|")
		RePage=(Arr_Rs(13,i)+1)\10
		If RePage<(Arr_Rs(13,i)+1)/10 Then RePage=RePage+1
		RePageUrl="topic.asp?id="&Arr_Rs(0,i)&"&boardid="&Arr_Rs(7,i)&"&tb="&Arr_Rs(14,i)
		Leftn=60
			If RePage>4 Then leftn=56
			If Repage>10 Then leftn=50
		If BBS.Fun.strLength(Arr_Rs(2,i)) > leftn Then
		  Caption=BBS.Fun.StrLeft(Arr_Rs(2,i),leftn-16) & BBS.Fun.StrRight(Arr_Rs(2,i),16) & "[长]"
		Else
		  Caption=Arr_Rs(2,i)
		End If
		S=Arr_Rs(16,i)
		If Not isNull(S) And S<>"" Then
			S=Split(S,"|")
			If S(0)<>"" Then Caption="<"&S(0)&">"&Caption&"</"&S(0)&">"
			If S(1)<>"" Then Caption="<span style='color:"&S(1)&"'>"&Caption&"</span>"
		End If
		'打开方式
		If BBS.Info(69)="1" Then S="target='_blank' " Else S=""
		Caption=UploadType&"<a "&S&"href='"&Repageurl&"'>"&Caption&"</a>"
		If Repage>1 Then
			Caption=Caption&" [<img src='images/Icon/gopage.gif' width=10 height=12> "
			If RePage<=5 Then
				For ii=2 To RePage
					Caption=Caption&"<a href='"&RePageurl&"&page="&ii&"'>"&ii&"</a> "
				Next
			Else
				For ii=2 To 4
					Caption=Caption&"<a href='"&RePageurl&"&page="&ii&"'>"&ii&"</a> "
				Next
					Caption=Caption&"... <a href='"&RePageurl&"&page="&RePage&"'>"&RePage&"</a> "
			End If
			Caption=Caption&" ]"
		End If
		S=BBS.ReadSkins("显示主题列表")
		If Datediff("n",Arr_Rs(8,i),BBS.NowbbsTime)<=180 Then Caption=Caption&BBS.SkinsPIC(18)
		S=Replace(S,"{状态}",Moodpic)		
		S=Replace(S,"{表情}",Facepic)
		S=Replace(S,"{用户名称}","<a href='userinfo.asp?name="&Arr_Rs(3,i)&"' title='查看 "&Arr_Rs(3,i)&" 的资料'>"&Arr_Rs(3,i)&"</a>")
		S=Replace(S,"{回复数}",Arr_Rs(13,i))
		S=Replace(S,"{点击数}",Arr_Rs(9,i))
		S=Replace(S,"{主题时间}",Arr_Rs(6,i))
		S=Replace(S,"{最后时间}",Arr_Rs(8,i))
		S=Replace(S,"{回复用户名称}","<a href='userinfo.asp?name="&LastRe(0)&"' title='查看 "&LastRe(0)&" 的资料'>"&LastRe(0)&"</a>")
		S=Replace(S,"{主题}",Caption)
		If Arr_Rs(4,i) >= 3 Then TopicLine = 0
		If Arr_Rs(4,i) < 3 and TopicLine = 0 Then
	      S = "<div class=topTopic>普通主题</div>" & S
		  TopicLine = 1
		End If
		TopicS=TopicS&S
	Next
	End If
	ShowTopic=TopicS
End Function
%>