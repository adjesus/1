<!--#include file="inc.asp"-->
<!--#include file="inc/Page_Cls.asp"-->
<%
Dim Key,Flag
If Not BBS.Founduser Then BBS.GotoErr(10)
If SESSION(CacheName& "MyGradeInfo")(20)="0" Then BBS.GotoErr(75)
If BBS.BoardID >0 Then BBS.CheckBoard()
BBS.Head"Search.asp",BBS.Boardname,"��̳����"
Key=BBS.Fun.Getkey("Key")
Flag=BBS.CheckNum(Request.querystring("Flag"))

If Key<>"" or Flag<>0 Then
	SearchList()
Else
	Main()
End If
BBS.Footer()
Set BBS =Nothing

Sub SearchList()
	with BBS
	Dim Temp,intPageNow,arr_Rs,i,Pages,Conut,p,PageInfo,Title,Sqlwhere,orders
	Dim Topic,TopicS,Caption,Moodpic,LastRe,RePageUrl,UploadType,RePage,leftn,ii,Stype,STime,again
	SType=.CheckNum(Request.querystring("SType"))
	STime=.CheckNum(Request.querystring("STime"))
	again=Request.querystring("again")
	Sqlwhere="IsDel=0 "
	Select Case Flag
	Case 1
		Title="��������"
		SqlWhere=Sqlwhere&"And IsGood=1"
	Case 2
		Title="��������"
		SqlWhere=Sqlwhere&"And DATEDIFF('d',[LastTime],'"&BBS.NowBbsTime&"')<1"
	Case 3
		Title="������������"
		Orders="hits desc,"
	Case 4
		Title="�����ظ�����"
		Orders="ReplyNum desc,"
	Case Else
		Title="ȫ������"
	End Select

	If Key<>"" Then
	If Len(Key)<2 Then .GotoErr(63)
	If again<>"" Then
		If Key<>.Fun.Getkey("Key1") Then Key=.Fun.Getkey("Key1")&" "&Key
	End If
	Select Case Stype
		Case"1":Sqlwhere=Sqlwhere&" And "&.Fun.SplitKey("Name",Key,"and")
		Case"2":Sqlwhere=Sqlwhere&" And "&.Fun.SplitKey("Caption",Key,"and")
		Case"3":Sqlwhere=Sqlwhere&" And "&"("&.Fun.SplitKey("Name",Key,"And")&" or "&.Fun.SplitKey("Caption",Key,"And")&")"
		Case Else
		.GotoErr(1)
	End Select
	Title="��̳���� �ؼ��֣�"&Key
	End If
	If STime<>0 Then Sqlwhere=Sqlwhere&" And DATEDIFF('d',[AddTime],'"&.NowBbsTime&"')<"&STime
	If .BoardID >0 Then Sqlwhere=sqlwhere&" And (BoardID="&.BoardID&" or TopType=5)"
	If .MyAdmin<>9 Then
	'����������������
		Temp=.NoShowTopic()
		If Temp<>"" Then Sqlwhere=Sqlwhere&" And BoardID not in ("&Temp&")"
		If Session(CacheName&"Searh")="" Then Session(CacheName&"Searh")=Sqlwhere
		If Session(CacheName&"Searh")<>Sqlwhere Then
			If (Session(CacheName&"SearchTime")+Int(.Info(17))/86400)>Now() Then .GoToErr(64)
			Session(CacheName&"Searh")=Sqlwhere
			Session(CacheName&"SearchTime")=Now()
		End If
	End If
	intPageNow = Request.QueryString("page")
	Set p = New Cls_PageView
	p.strTableName = "[Topic]"
	p.strPageUrl = "?flag="&Flag&"&Key="&Key&"&SType="&SType&"&TB="&.TB&"&STime="&STime&"&BoardID="&.BoardID
	P.strFieldsList = "TopicID,Face,Caption,Name,TopType,IsGood,AddTime,BoardID,LastTime,Hits,LastReply,UploadType,IsVote,ReplyNum,SqlTableID,IsLock,Font"
	p.strCondiction = SqlWhere
	p.strOrderList = Orders&"TopicID desc"
	p.strPrimaryKey = "TopicID"
	p.intPageSize = 20
	p.intPageNow = intPageNow
	p.strCookiesName = "Search"&SType&STime&.BoardID&.TB
	p.Reloadtime=0
	p.strPageVar = "page"
	p.InitClass
	Arr_Rs = p.arrRecordInfo
	PageInfo = p.strPageInfo
	Set p = nothing
		If IsArray(Arr_Rs) Then
	For i = 0 to UBound(Arr_Rs, 2)
		Moodpic=.SkinsPIC(16)
		If Arr_Rs(13,i) > Int(.Info(62)) Then Moodpic=.SkinsPIC(15)
		If Arr_Rs(5,i)=1 Then Moodpic=.SkinsPIC(13)'����
		If Arr_Rs(15,i)=1 Then Moodpic=.SkinsPIC(17)'����
		If Arr_Rs(12,i)=1 Then Moodpic=.SkinsPIC(14)'ͶƱ
		If Arr_Rs(4,i)=5 Then Moodpic=.SkinsPIC(10)'�ܶ�
		If Arr_Rs(4,i)=4 Then Moodpic=.SkinsPIC(11)'����
		If Arr_Rs(4,i)=3 Then Moodpic=.SkinsPIC(12)'��
		UploadType=""
		If Arr_Rs(11,i)<>"" Then Uploadtype="<img src='pic/FileType/"&Arr_Rs(11,i)&".gif' border='0' atl='"&Arr_Rs(11,i)&"' /> "
		LastRe=split(Arr_Rs(10,i),"|")
		RePage=(Arr_Rs(13,i)+1)\10
		If RePage<(Arr_Rs(13,i)+1)/10 Then RePage=RePage+1
		RePageUrl="topic.asp?id="&Arr_Rs(0,i)&"&boardid="&Arr_Rs(7,i)&"&tb="&Arr_Rs(14,i)&""
		Caption=.Fun.ReplaceKey(.Fun.StrLeft(Arr_Rs(2,i),60),Key)
		Temp=Arr_Rs(16,i)
		If Not isNull(Temp) And Temp<>"" Then
			Temp=Split(Temp,"|")
			If Temp(0)<>"" Then Caption="<"&Temp(0)&">"&Caption&"</"&Temp(0)&">"
			If Temp(1)<>"" Then Caption="<span style='color:"&Temp(1)&"'>"&Caption&"</span>"
		End If
		'�򿪷�ʽ
		If .Info(69)="1" Then Temp="target='_blank' " Else Temp=""
		Caption=UploadType&"<a "&Temp&"href='"&Repageurl&"'><span title='���ظ����ݣ�"&LastRe(1)&"'>"&Caption&"</span></a>"
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
		If Datediff("n",Arr_Rs(8,i),.NowbbsTime)<=180 Then Caption=Caption&.SkinsPIC(18)
		Caption=Moodpic&"<img src='pic/face/"&Arr_Rs(1,i)&".gIf' />"&Caption
		Topic="<div style='padding: 5px;text-align:left;border-bottom:1px solid "&.SkinsPIC(0)&";color:#5D7790'>"&.GetBoardName(arr_Rs(7,i))&"<br /> "&Caption&""&_
		"<div>���ߣ�"&Arr_Rs(3,i)& " | ����ʱ�䣺"&Arr_Rs(6,i)&" | �����"&Arr_Rs(9,i)&" | �ظ���"&Arr_Rs(13,i)&"</div></div>"
		TopicS=TopicS&Topic
	Next
	Topics=Topics&"<div style=""height:25px;BACKGROUND: "&.SkinsPIC(2)&";"">"&PageInfo&"</div>"
	TopMain()
	.ShowTable Title,TopicS
	Else
	.ShowTable "��̳����","<div style=""margin:18px;line-height:150%"">�Ҳ������������ݣ�<a href='javascript:history.go(-1)'>[����]</a></li></div>"
	End If
	End with
End Sub

Sub Main()
	Dim S
	S="<form method='get' style='margin:0'>"
	S=S&BBS.Row1("<div style=""padding:4px""><b>����˵����</b><li>����̳ÿ�������ļ��ʱ��Ϊ"&BBS.Info(17)&"��</li><li>���Բ��÷ִ�������������</li></div>")
	S=S&BBS.Row("<b>�����ؼ��֣�</b>","<input type='text' name='Key' size='52' class='text'>","65%","")
	S=S&BBS.Row("<b>�������ͣ�</b>","<input type='radio' value='1' name='SType'> ���������� <input type='radio' name='SType' checked value='2'> ���������� <input type='radio' value='3' name='SType'>���߾���","65%","")
	S=S&BBS.Row("<b>�������ڷ�Χ��</b>","<select size='1' name='STime'><option selected value='0'>��������</option><option value='1'>1������</option><option value='2'>2������</option><option value='7'>7������</option><option value='15'>15������</option><option value='30'>30������</option></select>","65%","")
	S=S&BBS.Row("<b>��������̳��</b>","<select name='BoardID'><option value='0'>����ȫ����̳</option>"& BBS.BoardIDList(0,0)&"</select>","65%","")
	S=S&"<div style=""padding:2px;BACKGROUND: "&BBS.SkinsPIC(2)&";"" align=""center""><input class='button' type=""submit"" value=""��ʼ����"" /></div></form>"
	BBS.ShowTable "��̳����",S
End Sub
Sub topMain()
	Dim S
	S="<form method='get' style='margin:0'>"
	S=S&BBS.Row1("<div style=""padding:4px""><b>����˵����</b><li>����̳ÿ�������ļ��ʱ��Ϊ"&BBS.Info(17)&"��</li><li>���Բ��÷ִ�������������</li></div>")
	S=S&BBS.Row("<b>�����ؼ��֣�</b>","<input type='text' name='Key' size='50' class='text' value='"&BBS.Fun.HtmlCode(Key)&"' /><input type='hidden' name='Key1' value='"&BBS.Fun.HtmlCode(Key)&"' />","65%","")
	S=S&BBS.Row("<b>�������ͣ�</b>","<input type='radio' value='1' name='SType'> ���������� <input type='radio' name='SType' checked value='2'> ���������� <input type='radio' value='3' name='SType'>���߾���","65%","")
	S=S&BBS.Row("<b>�������ڷ�Χ��</b>","<select size='1' name='STime'><option selected value='0'>��������</option><option value='1'>1������</option><option value='2'>2������</option><option value='7'>7������</option><option value='15'>15������</option><option value='30'>30������</option></select>","65%","")
	S=S&BBS.Row("<b>��������̳��</b>","<select name='BoardID'><option value='0'>����ȫ����̳</option>"& BBS.BoardIDList(BBS.BoardID,0)&"</select>","65%","")
	S=S&"<div style=""padding:2px;BACKGROUND: "&BBS.SkinsPIC(2)&";"" align=""center""><input class='button' type=""submit"" value=""��ʼ����"" /> <input class='button' name='again' type=""submit"" value=""�ڽ������"" /></div></form>"
	BBS.ShowTable "��̳����",S
End Sub
%>
