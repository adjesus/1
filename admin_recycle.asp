<!--#include file="Admin_Check.asp"-->
<!--#include file="inc/page_Cls.asp"-->
<!--#include file="inc/Ubb_Cls.asp"-->
<%
Dim Action,username,bbsID,ID,PageInfo
CheckString "36"
username=BBS.MyName
Head()
Action=request.querystring("Action")
Select Case Action
Case "Submit"
	Submit()
Case "See"
	See()
Case "Del"
	Del()
Case "TBInfo"
	TBInfo()
Case "Giveback"
	Giveback()
Case "DelAll"
	DelAll()
Case Else
	Recycle()
End Select
Footer()
Set BBS =Nothing

Function GetPageInfo(PTable,PFieldslist,PCondiction,POrderlist,PPrimaryKey,PSize,PCookiesName,Purl)
	Dim P
	Set P = New Cls_PageView
	P.strTableName =PTable
	P.strFieldsList =PFieldslist
	P.strCondiction =PCondiction
	P.strOrderList = POrderlist
	P.strPrimaryKey = PPrimaryKey
	P.intPageSize = PSize
	P.intPageNow = Request("page")
	P.strCookiesName = PCookiesName
	P.strPageUrl = PUrl
	P.InitClass
	GetPageInfo = P.arrRecordInfo
	PageInfo = P.strPageInfo
	Set P = nothing
End Function

Sub Recycle()
Response.Write"<form name='kk'  style='margin:0' method='POST' action='?Action=Submit'>"&_
"<div class='mian'><div class='top'>回收站</div><div class='divtr2' style='padding:3px'><div style='float:right'>【<a onclick=checkclick('您确定要清空回收站的全部帖子吗？','?Action=DelAll')>"&IconD&"清空回收站</a>】</div>【<a href='?action=Recycle'><font color='red'>列出全部主题</font></a>】"&TBList(0)&" </div>"
	Dim arr_Rs,i
	Dim Temp,BbsID	
	Arr_rs=GetPageInfo("[Topic]","TopicID,SqlTableID,Face,Caption,Name,LastTime,BoardID,ReplyNum","IsDel=1","TopicID desc","TopicID",20,"Recycle"&BBS.TB,"?action=Recycle")
	If IsArray(Arr_Rs) Then
Response.Write"<table class='Stable' border='0' cellpadding='3' cellspacing='0'><tr><th width='35'>选择</th><th width='55%'>帖子</th><th width='12%'>作者</th><th width='8%'>回复</th><th>最后时间</th></tr>"
For i = 0 to UBound(Arr_Rs, 2)
	Set Rs=BBS.Execute("Select BbsID From[Bbs"&Arr_Rs(1,i)&"] Where TopicID="&Arr_Rs(0,i)&" and BoardID="&Arr_Rs(6,i))
	If Not Rs.Eof Then BbsID=Rs(0)
	Rs.Close
	    Response.Write"<tr>"
		Response.Write"<td><input type='checkbox' name='Topic' value='"&Arr_Rs(0,i)&"|"&Arr_Rs(6,i)&"|"&Arr_Rs(1,i)&"'></td>"&_
		"<td><img src='pic/face/"&Arr_Rs(2,i)&".gIf' align='absmiddle'><a href=?Action=See&BbsID="&BbsID&"&TopicID="&Arr_Rs(0,i)&">"&BBS.Fun.StrLeft(Arr_Rs(3,i),35)&"</td>"&_
		"<td><a target=_blank  href='UserInfo.asp?name="&Arr_Rs(4,i)&"' title='查看 "&Arr_Rs(4,i)&" 的资料'>"&Arr_Rs(4,i)&"</a></td><td>"&Arr_Rs(7,i)&"</td>"&_
		"<td>"&Arr_Rs(5,i)&"</td></tr>"
	Next
	Response.Write"</table><div class='bottom'><input type='checkbox' name='chkall' value='on' onClick='CheckAll(this.form)' />全选&nbsp;<input type='submit' class='button' value='删除所选' name='Go'><input class='button' type='submit' value='还原所选' name='Go'></div><div class='divtr2'>"&PageInfo&"</div>"	
	Else
	Response.Write"<div class='divtr1'><br />&nbsp;没有发现删除的主题帖<br />&nbsp;</div>"
	End If
	Response.Write"</div></form>"
End Sub

Sub TBInfo()
Response.Write"<form name='kk'  style='margin:0' method='POST' action='?Action=Submit'>"&_
"<div class='mian'><div class='top'>回收站</div><div class='divtr2' style='padding:3px'><div style='float:right'>【<a href=#this onclick=checkclick('您确定要清空回收站的全部帖子吗？','?Action=DelAll')>"&IconD&"清空回收站</a>】</div>【<a href='?action=Recycle'>列出全部主题</a>】"&TBList(BBS.TB)&"</div>"
	Dim intPageNow,arr_Rs,i,Pages,Conut,page,strPageInfo
	Dim Temp
		Arr_rs=GetPageInfo("[BBS"&BBS.TB&"]","BbsID,TopicID,Face,Caption,Name,LastTime,ReplyTopicID,BoardID","IsDel=1","BbsID desc","BbsID",20,"Recycle"&BBS.TB,"?action=TBInfo")
	If IsArray(Arr_Rs) Then
	Response.Write"<table class='Stable' border='0' cellpadding='3' cellspacing='0'><tr><th width='35'>选择</th><th width='55%'>帖子</th><th width='18%'>作者</th><th>最后时间</th></tr>"
	For i = 0 to UBound(Arr_Rs, 2)
	Response.Write"<tr>"
	Response.Write"<td><input type='checkbox' "
	IF Arr_Rs(1,i)=0 Then
		Response.Write "name='Reply' value='"&Arr_Rs(0,i)&"|"&Arr_Rs(6,i)&"|"&Arr_Rs(7,i)&"|"&BBS.TB&"'"
	Else
		Response.Write "name='Topic' value='"&Arr_Rs(1,i)&"|"&Arr_Rs(7,i)&"|"&BBS.TB&"'"
	End If
	Response.Write"></td>"&_
		"<td><img src='pic/face/"&Arr_Rs(2,i)&".gIf' align='absmiddle'><a href='?Action=See&BbsID="&Arr_Rs(0,i)&"&TopicID="&Arr_Rs(1,i)&"'>"&BBS.Fun.StrLeft(Arr_Rs(3,i),25)&"</a></td>"&_
		"<td align='center'><a target=_blank  href='UserInfo.asp?name="&Arr_Rs(4,i)&"' title='查看 "&Arr_Rs(4,i)&" 的资料'>"&Arr_Rs(4,i)&"</a></td>"&_
		"<td align='center'>"&Arr_Rs(5,i)&"</td></tr>"
	Next
	Response.Write"</table>"
	Response.Write"<div class='bottom'><input type='checkbox' name='chkall' value='on' onClick='CheckAll(this.form)' />全选&nbsp;<input type='submit' class='button' value='删除所选' name='Go'><input type='submit' class='button' value='还原所选' name='Go'></div><div class='divtr2'>"&PageInfo&"</div>"	
	Else
	Response.Write"<div class='divtr1'><br />&nbsp;这个数据表中没有发现被删除的帖子<br />&nbsp;</div>"
	End If
	Response.Write"</div></form>"
End Sub


Sub Del()
	Dim BbsID,TopicID
	BbsID=Request.querystring("BbsID")
	TopicID=request.querystring("TopicID")
	If TopicID=0 then
	BBS.Execute("Delete From [Bbs"&BBS.TB&"] where IsDel=True And BbsID="&BbsID)
	BBS.Execute("Delete From [Appraise] where BbsID="&BbsID&" And TopicID="&TopicID)
	Suc"","成功删除了这个回复帖！","?"
	Else
	BBS.Execute("Delete From [Topic] where IsDel=True And TopicID="&TopicID)
	BBS.Execute("Delete From [TopicVote] where TopicID="&TopicID)
	BBS.Execute("Delete From [TopicVoteUser] where TopicID="&TopicID)
	BBS.Execute("Delete From [Bbs"&BBS.TB&"] where BbsID="&BbsID&" or ReplyTopicID="&TopicID)
	BBS.Execute("Delete From [Appraise] where TopicID="&TopicID)
	Suc"","成功删除这个主题（包括其回复帖）！","?"
	End if
End Sub

Sub DelAll()
	Dim AllTable,I
	AllTable=Split(BBS.BBStable(0),",")
	Set Rs=BBS.Execute("Select TopicID,SqlTableID From [Topic] where IsDel=1")
	Do while Not Rs.eof
		BBS.Execute("Delete * From [Bbs"&Rs(1)&"] where ReplyTopicID="&Rs(0)&"")
	Rs.movenext
	Loop
	Rs.Close
	For i=0 To uBound(AllTable)
		BBS.Execute("Delete * From [Bbs"&AllTable(i)&"] where IsDel=1")
	Next
	BBS.Execute("Delete From [Topic] where IsDel=1")
	BBS.execute("delete * from [TopicVote] where  not exists (select name from [Topic] where [TopicVote].TopicID=[Topic].TopicId)")
	BBS.execute("delete * from [TopicVoteUser] where  not exists (select name from [Topic] where [TopicVoteUser].TopicID=[Topic].TopicId)")
	BBS.execute("delete * from [Appraise] where  not exists (select TopicID from [Topic] where [Appraise].TopicID=[Topic].TopicId)")
	Suc"","成功清空了回收站！","?"
End Sub

Sub Giveback
	Dim BbsID,TopicID,ReplyTopicID,BoardID,Temp
	BbsID=request.querystring("BbsID")
	Set Rs=BBS.Execute("Select TopicID,ReplyTopicID,BoardID,IsDel From[Bbs"&BBS.TB&"] where BbsID="&BbsID)
	IF Rs.eof Then
		GoBack"","该帖不存在或者已经被永久删除":Exit Sub
	End IF
	If Rs(0)=0 And Rs(1)<>0 then
		BBS.Execute("Update [Config] Set AllEssayNum=AllEssayNum+1")
		BBS.Execute("Update [Board] Set EssayNum=EssayNum+1 Where BoardID="&Rs(2)&" And ParentID<>0")
		BBS.Execute("Update [Topic] Set ReplyNum=ReplyNum+1,IsDel=0 where TopicID="&Rs(1))
		BBS.Execute("Update [Bbs"&BBS.TB&"] Set IsDel=0 where TopicID="&Rs(1)&" or BbsID="&BbsID)
	Else
		Temp=BBS.Execute("Select ReplyNum From[Topic] where TopicID="&Rs(0))(0)
		BBS.Execute("Update [Config] Set TopicNum=TopicNum+1,AllEssayNum=AllEssayNum+"&Temp+1&"")
		BBS.Execute("Update [Board] Set EssayNum=EssayNum+"&Temp+1&",TopicNum=TopicNum+1 Where BoardID="&Rs(2)&" And ParentID<>0")
		BBS.Execute("Update [Topic] Set IsDel=0 where TopicID="&Rs(0))
		BBS.Execute("update [Bbs"&BBS.TB&"] Set IsDel=0 where BbsID="&BbsID)
	End if
	Rs.Close
	Suc"","成功的恢复帖子","?"
End Sub

Sub Submit()
Dim Topic,Reply,Go,Temp,i
Topic=Request.form("Topic")
Reply=Request.form("Reply")
IF Topic="" And Reply="" Then GoBack"","请先选择项目。":Exit Sub
Topic=split(Topic,",")
Reply=split(Reply,",")
Go=Request.form("Go")
	If Go="删除所选" then
		For i=0 to ubound(Topic)
		Temp=split(Topic(I),"|")
		BBS.Execute("Delete From [Bbs"&Temp(2)&"] where TopicID="&Temp(0)&" or ReplyTopicID="&Temp(0))
		BBS.Execute("Delete From [Topic] where TopicID="&Temp(0))
		BBS.Execute("Delete From [TopicVote] where TopicID="&Temp(0))
		BBS.Execute("Delete From [TopicVoteUser] where TopicID="&Temp(0))
		BBS.Execute("Delete From [Appraise] where TopicID="&Temp(0))
		Next
		For i=0 to ubound(Reply)
		Temp=split(Reply(I),"|")
		BBS.Execute("Delete From [Bbs"&Temp(3)&"] where BbsID="&Temp(0))
		BBS.Execute("Delete From [Appraise] where BbsID="&Temp(0)&" And TopicID="&Temp(1))
		Next
		Suc"","成功的删除所选的帖子","?"
	ElseIF Go="还原所选" then
		Dim TempNum
		For i=0 to ubound(Topic)
			Temp=split(Topic(I),"|")
			TempNum=BBS.Execute("Select ReplyNum From[Topic] where TopicID="&Temp(0))(0)
			BBS.Execute("Update [Config] Set TopicNum=TopicNum+1,AllEssayNum=AllEssayNum+"&TempNum+1&"")
			BBS.Execute("Update [Board] Set EssayNum=EssayNum+"&TempNum+1&",TopicNum=TopicNum+1 Where BoardID="&Temp(1)&" And ParentID<>0")
			BBS.Execute("Update [Topic] Set IsDel=0 where TopicID="&Temp(0))
			BBS.Execute("update [Bbs"&Temp(2)&"] Set IsDel=0 where TopicID="&Temp(0))
		Next
		For i=0 to ubound(Reply)
		Temp=split(Reply(I),"|")
		Set Rs=BBS.Execute("Select Top 1 BbsID From[Bbs"&Temp(3)&"] where BbsID="&Temp(0)&" And IsDel=1")
		If Not Rs.Eof Then
		BBS.Execute("Update [Config] Set AllEssayNum=AllEssayNum+1")
		BBS.Execute("Update [Board] Set EssayNum=EssayNum+1 Where BoardID="&Temp(2)&" And ParentID<>0")
		BBS.Execute("Update [Topic] Set ReplyNum=ReplyNum+1,IsDel=0 where TopicID="&Temp(1))
		BBS.Execute("Update [Bbs"&Temp(3)&"] Set IsDel=0 where TopicID="&Temp(1)&" or BbsID="&Temp(0))
		End If
		Rs.Close
		Next
		Suc"","成功的还原所选的帖子","?"
	End If
End SUB


Sub See()
Dim BbsID,IUBB,EssayType,TopicID,ReplyTopicID,Arr_Rs,i,Sqlwhere
BbsID=Trim(Request.querystring("BbsID"))
TopicID=Request.querystring("TopicID")
ReplyTopicID=Request.querystring("ReplyTopicID")
If ReplyTopicID="" Then Sqlwhere="TopicID="&TopicID&" or ReplyTopicID="&TopicID
If TopicID="" or TopicID="0" Then Sqlwhere="BBSID="&BBSID

Arr_rs=GetPageInfo("[Bbs"&BBS.TB&"]","BbsID,Caption,Content,Name,LastTime,BoardID,TopicID,ReplyTopicID,UbbString,Face,IP",Sqlwhere,"TopicID desc","BbsID",10,"Recycle"&BBSID,"?action=See&BBSID="&BBSID&"&TopicID="&TopicID&"&ReplyTopicID="&ReplyTopicID)
If IsArray(Arr_Rs) Then
Response.Write"<div class='mian'><div class='top'><a style='float:right;color:#FFF' href='javascript:history.go(-1)'>【返回】</a>查看帖子</div>"
Set IUBB=New Cls_IUBB
For i = 0 to UBound(Arr_Rs, 2)
IUBB.UbbString=Arr_Rs(8,i)
ID=Arr_Rs(6,i)
If Arr_Rs(6,i)<>0 Then
	EssayType="主题帖："
	TopicID=Arr_Rs(6,i)
Response.Write"<div class='divtr2'>"
Else
	EssayType="回复帖："
Response.Write"<div class='divtr1'>"
End If
Response.Write"<div style=""min-height:150px;font-size:9pt;line-height:normal;padding:5px;word-wrap : break-word ;word-break : break-all ;"" onload=""this.style.overflowX='auto';"">"&EssayType&BBS.Fun.HtmlCode(Arr_Rs(1,i))&"<hr size='1' color=#DCE2E4>"
Response.Write" <blockquote>"
Response.Write"<img src='pic/face/"&Arr_Rs(9,i)&".gIf' align='absmiddle'>"
If Arr_Rs(7,i)=0 Then Response.Write "<b>"&BBS.Fun.HtmlCode(Arr_Rs(1,i))&"</b>"
Response.Write"<br>"&IUBB.UBB(Arr_Rs(2,i),1)&"</Span></blockquote><hr size='1' color='#DCE2E4'>"&_
"<div style='FLOAT: right;'><a href='?Action=Del&BbsID="&Arr_Rs(0,i)&"&TopicID="&Arr_Rs(6,i)&"&TB="&BBS.TB&"'>"&IconD&"永久删除</a>"
If i=0 Then
Response.Write" <a href='?Action=Giveback&BbsID="&Arr_Rs(0,i)&"&TB="&BBS.TB&"&BoardID="&Arr_Rs(5,i)&"'><img src='Images/icon/giveback.gif' border='0' align='absmiddle'> 还原帖子</a>"
End If
Response.Write"</div>作者：<a href='Admin_user.asp?Action=EditUser&ID="&Arr_Rs(3,i)&"'>"&IconE&Arr_Rs(3,i)&"</a>&nbsp;&nbsp;IP：<a href='Admin_Action.asp?action=AddLockIp&IP="&Arr_Rs(10,I)&"&Readme=用户名："&Arr_Rs(3,I)&"' tilte='封锁用户IP'><img src='Images/icon/lock.gif' border='0' alt='封锁IP'  align=""absmiddle"" /> "&Arr_Rs(10,i)&"</a>&nbsp;&nbsp;更新时间："&Arr_Rs(4,i)&"</div></div>"
Next
Set IUBB=Nothing
Response.Write "<div class='divtr2'>"&PageInfo&"</div></div>"
Else
	GoBack"","该帖不存在或者已经被永久删除"
End If
End Sub


Function TBList(Num)
	Dim AllTable,I,Temp
	AllTable=Split(BBS.BBStable(0),",")
	For i=0 To uBound(AllTable)
		If Int(AllTable(i))=Int(Num) Then
		Temp=Temp&"【<font color=red>数据表"&AllTable(i)&"</font>】"
		Else
		Temp=Temp&"【<a href='?Action=TBInfo&TB="&AllTable(i)&"'>数据表"&AllTable(i)&"</a>】"
		End IF
	next
	TBList=Temp
End Function

%>

<script language="JavaScript">
<!--
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall'){
	e.checked = form.chkall.checked;
	}
   }
  }
//-->
</script>