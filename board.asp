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
	area=.ReadSkins("���������")
	area2=.ReadSkins("�����������")
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
				BoardStr=Replace(BigBoard,"{��������}",.Board_Rs(3,i))
				BoardStr=Replace(BoardStr,"{����ID}",.Board_Rs(1,i))
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
	BoardStr=Replace(BoardStr,"{��ʾ���}",BoardChild)
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
	S=BBS.ReadSkins("�����б���")
	S=Replace(S,"{������ť}",Button)
	S=Replace(S,"{����}",BBS.Boardadmin)
	S=Replace(S,"{���ID}",BBS.boardid)
	S=Replace(S,"{���ݱ�ID}",BBS.tb)
	S=Replace(S,"{����}",BBS.Placard(BBS.boardid))
	S=Replace(S,"{�������}",BBS.Boardname)
	If BBS.Info(21)="1" then S=Replace(S,"{��ʾ����}",ShowOnline()) Else S=Replace(S,"{��ʾ����}","")
	S=Replace(S,"{��������б�}","<script language=""JavaScript"" type=""text/javascript"">BoardSelect()</script>")
	S1=ShowTopic()
	S=Replace(S,"{��ҳ}",strPageInfo)
	S=Replace(S,"{��ʾ�����б�}",S1)
	ShowTopicList=S
End Function

Function Showonline()
	Dim S,Boardupdate
	S=BBS.ReadSkins("��ʾ�������")
	S=Replace(S,"{��ʾ�����б�}","<div id=""showon""></div>")
	If Session(CacheName&"Online")="1" Then S=Replace(S,"{չ�������б�ͼ��}","-.gif") Else S=Replace(S,"{չ�������б�ͼ��}","+.gif")	
	S=Replace(S,"{���ID}",BBS.boardid)
	S=Replace(S,"{��������}",BBS.AllOnlineNum)
	S=Replace(S,"{������������}",BBS.BoardOnlineNum)
	S=Replace(S,"{�������߻�Ա��}",BBS.BoardUserOnlineNum)
	S=Replace(S,"{���������ο���}",Int(BBS.BoardOnlineNum)-int(BBS.BoardUserOnlineNum))
	S=Replace(S,"{����Ա}",BBS.SkinsPic(21))
	S=Replace(S,"{��������}",BBS.SkinsPic(22))
	S=Replace(S,"{����}",BBS.SkinsPic(23))
	S=Replace(S,"{VIP��Ա}",BBS.SkinsPic(24))
	S=Replace(S,"{��Ա}",BBS.SkinsPic(25))
	S=Replace(S,"{����}",BBS.SkinsPic(26))
	S=Replace(S,"{�ο�}",BBS.SkinsPic(27))
	Boardupdate=BBS.GetEachBoardCache(BBS.boardid)
	S=Replace(S,"{��������}",Boardupdate(2))
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
	P.strCookiesName = "List"&BBS.boardid&"BBS"&Flag'�ͻ��˼�¼����
	'P.Reloadtime=3'Ĭ�������Ӹ���Cookies
	P.InitClass
	Arr_Rs = P.arrRecordInfo
	strPageInfo = P.strPageInfo
	Set P = nothing
	If IsArray(Arr_Rs) Then
	TopicLine = -1
	For i = 0 to UBound(Arr_Rs, 2)
		Moodpic=BBS.SkinsPIC(16)
		If Arr_Rs(13,i) > Int(BBS.Info(62)) Then Moodpic=BBS.SkinsPIC(15)
		If Arr_Rs(5,i)=1 Then Moodpic=BBS.SkinsPIC(13)'����
		If Arr_Rs(15,i)=1 Then Moodpic=BBS.SkinsPIC(17)'����
		If Arr_Rs(12,i)=1 Then Moodpic=BBS.SkinsPIC(14)'ͶƱ
		If Arr_Rs(4,i)=5 Then Moodpic=BBS.SkinsPIC(10)'�ܶ�
		If Arr_Rs(4,i)=4 Then Moodpic=BBS.SkinsPIC(11)'����
		If Arr_Rs(4,i)=3 Then Moodpic=BBS.SkinsPIC(12)'��
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
		  Caption=BBS.Fun.StrLeft(Arr_Rs(2,i),leftn-16) & BBS.Fun.StrRight(Arr_Rs(2,i),16) & "[��]"
		Else
		  Caption=Arr_Rs(2,i)
		End If
		S=Arr_Rs(16,i)
		If Not isNull(S) And S<>"" Then
			S=Split(S,"|")
			If S(0)<>"" Then Caption="<"&S(0)&">"&Caption&"</"&S(0)&">"
			If S(1)<>"" Then Caption="<span style='color:"&S(1)&"'>"&Caption&"</span>"
		End If
		'�򿪷�ʽ
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
		S=BBS.ReadSkins("��ʾ�����б�")
		If Datediff("n",Arr_Rs(8,i),BBS.NowbbsTime)<=180 Then Caption=Caption&BBS.SkinsPIC(18)
		S=Replace(S,"{״̬}",Moodpic)		
		S=Replace(S,"{����}",Facepic)
		S=Replace(S,"{�û�����}","<a href='userinfo.asp?name="&Arr_Rs(3,i)&"' title='�鿴 "&Arr_Rs(3,i)&" ������'>"&Arr_Rs(3,i)&"</a>")
		S=Replace(S,"{�ظ���}",Arr_Rs(13,i))
		S=Replace(S,"{�����}",Arr_Rs(9,i))
		S=Replace(S,"{����ʱ��}",Arr_Rs(6,i))
		S=Replace(S,"{���ʱ��}",Arr_Rs(8,i))
		S=Replace(S,"{�ظ��û�����}","<a href='userinfo.asp?name="&LastRe(0)&"' title='�鿴 "&LastRe(0)&" ������'>"&LastRe(0)&"</a>")
		S=Replace(S,"{����}",Caption)
		If Arr_Rs(4,i) >= 3 Then TopicLine = 0
		If Arr_Rs(4,i) < 3 and TopicLine = 0 Then
	      S = "<div class=topTopic>��ͨ����</div>" & S
		  TopicLine = 1
		End If
		TopicS=TopicS&S
	Next
	End If
	ShowTopic=TopicS
End Function
%>