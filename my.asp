<!--#include file="Inc.asp"-->
<!--#include file="Inc/page_Cls.asp"-->
<%
Dim TextInfo,action,Title
Dim SqlWhere
Dim Caption,Content,tPageUrl

If Not BBS.Founduser Then BBS.GoToErr(10)
If BBS.BoardID>0 Then
	BBS.CheckBoard()
	SqlWhere="BoardID="&BBS.BoardID&" And IsDel=0 And "
	TextInfo="查看"&BBS.BoardName
Else
	SqlWhere="IsDel=0 And "
	BBS.Position=BBS.Position&" -> <a href='userinfo.asp'>用户控制面版<a>"
	TextInfo="查看"
End If
action=Lcase(Request.querystring("action"))
If len(action)>10 Then BBS.GotoErr(1)
Select Case action
Case"mysay"
	Title="本人帖子列表"
	BBS.Head "","",TextInfo&Title
	SqlWhere=SqlWhere&"Name='"&BBS.MyName&"'"
Case"myreply"
	Title="本人回复列表"
	BBS.Head "","",TextInfo&Title
	SqlWhere=SqlWhere&"Name<>'"&BBS.MyName&"' and (TopicID in (select ReplyTopicID from [bbs"&BBS.TB&"] where name='"&BBS.MyName&"' And IsDel=0))"
Case"mygood"
	Title="本人精华帖子列表"
	BBS.Head "","",TextInfo&Title
	SqlWhere=SqlWhere&"name='"&BBS.MyName&"' and IsGood"
Case"last"
	Title="上次访问后论坛的新帖"
	BBS.Head "","",TextInfo&Title
	SqlWhere =SqlWhere&"Datediff('s',LastTime,'"&BBS.GetMemor("","LastTime")&"')<0"
Case Else
	BBS.GotoErr(1)
End Select
If BBS.BoardID=0 Then
	MyManager()
	tPageUrl="?action="&action&"&boardid="&BBS.BoardID
Else
 	tPageUrl="?action="&action
End If
ShowTopic()
BBS.Footer()
Set BBS =Nothing



Function ShowTopic()
	Dim Temp,intPageNow,arr_Rs,i,Pages,Conut,p,PageInfo
	Dim Topic,TopicS,Caption,Moodpic,LastRe,RePageUrl,UploadType,RePage,leftn,ii
	Set P = New Cls_PageView
	P.strTableName = "[Topic]"
	P.strPageUrl = tPageurl
	P.strFieldsList = "TopicID,Face,Caption,Name,TopType,IsGood,AddTime,BoardID,LastTime,Hits,LastReply,UploadType,IsVote,ReplyNum,SqlTableID,IsLock,Font"
	P.strCondiction = SqlWhere
	P.strOrderList = "TopType desc,LastTime desc"
	P.strPrimaryKey = "TopicID"
	P.intPageSize = 15
	P.intPageNow=request("page")
	P.strCookiesName = action
	P.Reloadtime=10
	P.InitClass
	Arr_Rs = P.arrRecordInfo
	PageInfo = P.strPageInfo
	Set P = nothing
	If IsArray(Arr_Rs) Then
	For i = 0 to UBound(Arr_Rs, 2)
		Moodpic=BBS.SkinsPIC(16)
		If Arr_Rs(13,i) > Int(BBS.Info(62)) Then Moodpic=BBS.SkinsPIC(15)
		If Arr_Rs(5,i)=1 Then Moodpic=BBS.SkinsPIC(13)'精华
		If Arr_Rs(15,i)=1 Then Moodpic=BBS.SkinsPIC(17)'锁定
		If Arr_Rs(12,i)=1 Then Moodpic=BBS.SkinsPIC(14)'投票
		If Arr_Rs(4,i)=5 Then Moodpic=BBS.SkinsPIC(10)'总顶
		If Arr_Rs(4,i)=4 Then Moodpic=BBS.SkinsPIC(11)'区顶
		If Arr_Rs(4,i)=3 Then Moodpic=BBS.SkinsPIC(12)'顶
		UploadType=""
		If Arr_Rs(11,i)<>"" Then Uploadtype="<img src='pic/FileType/"&Arr_Rs(11,i)&".gif' border='0' atl='"&Arr_Rs(11,i)&"' /> "
		LastRe=split(Arr_Rs(10,i),"|")
		RePage=(Arr_Rs(13,i)+1)\10
		If RePage<(Arr_Rs(13,i)+1)/10 Then RePage=RePage+1
		RePageUrl="topic.asp?id="&Arr_Rs(0,i)&"&boardid="&Arr_Rs(7,i)&"&tb="&Arr_Rs(14,i)&""
		
		Caption=BBS.Fun.StrLeft(Arr_Rs(2,i),60)
		Temp=Arr_Rs(16,i)
		If Not isNull(Temp) And Temp<>"" Then
			Temp=Split(Temp,"|")
			If Temp(0)<>"" Then Caption="<"&Temp(0)&">"&Caption&"</"&Temp(0)&">"
			If Temp(1)<>"" Then Caption="<span style='color:"&Temp(1)&"'>"&Caption&"</span>"
		End If
		'打开方式
		If BBS.Info(69)="1" Then Temp="target='_blank' " Else Temp=""
		Caption=UploadType&"<a "&Temp&"href='"&Repageurl&"'><span title='最后回复内容："&LastRe(1)&"'>"&Caption&"</span></a>"
		If Repage>1 Then
			Caption=Caption&" [<img src='images/Icon/gopage.gif' width='10' height='12' /> "
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
		If Datediff("n",Arr_Rs(8,i),BBS.NowbbsTime)<=180 Then Caption=Caption&BBS.SkinsPIC(18)
		Caption=Moodpic&"<img src='pic/face/"&Arr_Rs(1,i)&".gIf' />"&Caption
		Topic="<div style='padding: 5px;text-align:left;border-bottom:1px solid "&BBS.SkinsPIC(0)&";color:#5D7790'>"&BBS.GetBoardName(arr_Rs(7,i))&"<br /> "&Caption&""&_
		"<div>作者："&Arr_Rs(3,i)& " | 发表时间："&Arr_Rs(6,i)&" | 浏览："&Arr_Rs(9,i)&" | 回复："&Arr_Rs(13,i)&"</div></div>"
		TopicS=TopicS&Topic
	Next
	Topics=Topics&"<div style=""height:25px;BACKGROUND: "&BBS.SkinsPIC(2)&";"">"&PageInfo&"</div>"
	End If
	BBS.ShowTable Title,TopicS
	
End Function

Sub MyManager()
	Response.Write BBS.ReadSkins("用户控制面版")
End Sub
%>