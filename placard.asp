<!--#include file="Inc.asp"-->
<!--#include file="Inc/Page_Cls.asp"-->
<%
Dim Action
If Not BBS.Founduser Then BBS.GoToerr(10)
IF SESSION(CacheName& "MyGradeInfo")(37)="0" Then BBS.GoToErr(67)
Action=Lcase(Request.querystring("Action"))
If Len(Action)>10 then BBS.GoToerr(1)
Select Case Action
Case"del"
	DelPlacard
Case Else
	Placard()
End Select
BBS.Footer()
Set BBS =Nothing

Sub DelPlacard()
	Dim ID
	ID=BBS.Checknum(request.querystring("ID"))
	IF BBS.MyAdmin=7 Then
		BBS.execute("Delete From [Placard] where ID="&ID&" and Name='"&BBS.MyName&"'")
	Else
		BBS.execute("Delete From [Placard] where ID="&ID&"")
	End IF
	BBS.Cache.clean("Placard")
	BBS.NetLog "删除公告"
	Response.redirect "Placard.asp"
End Sub

Sub Placard()
	Dim Caption,Content,Temp,TmpBoardID,S,Title,Rs,ID
	If BBS.BoardID>0 Then BBS.CheckBoard()
	IF BBS.MyAdmin=7 And BBS.IsBoardAdmin=False  Then BBS.GoToErr(68)
	ID=BBS.Checknum(request("ID"))
	BBS.Head"Placard.asp?Action=Say&BoardID="&BBS.BoardID,BBS.BoardName,"发布公告"
	Caption=BBS.Fun.Checkbad(BBS.Fun.GetStr("caption"))
	Content=BBS.Fun.Checkbad(BBS.Fun.GetStr("Content"))
	IF Caption="" And Content="" Then
		Title="发布公告"
		TmpBoardID=BBS.BoardID
		If ID<>0 Then
			Set Rs=BBS.execute("select BoardID,Caption,Content,AddTime,Name,hits from [Placard] where ID="&ID&"")
			IF Not rs.eof Then
				Title="编辑公告"
				TmpBoardID=Rs(0)
				Caption=Rs(1)
				Content=Rs(2)
				If BBS.MyAdmin=7 Then
					If Lcase(BBS.MyName)<>Lcase(Rs(4)) Then BBS.GotoErr(73)
				End If
			Else
				BBS.Gotoerr(69)
			End If
		End If
		S="<form  style='margin:0' action='?action=Placard&BoardID="&BBS.BoardID&"' method='post' name='say'><textarea id='content' name='content' style='display:none'>"&Content&"</textarea>"
		If BBS.MyAdmin=7 Then
			Temp=BBS.BoardName&"<input name='BoardID' value='"&BBS.BoardID&"' type='hidden' />"
		Else
			Temp="<select name='BoardID'><option value='0'>论坛首页</option>"&BBS.BoardIDList(TmpBoardID,0)&"</select>"
		End If
		S=S&BBS.Row("<b>公告标题：</b>","<input name='ID' type='hidden' value='"&ID&"' /><input type=hidden name='iCode' id='iCode' value='BBS' /><input name='caption' type='text' class='text' id='caption' value='"&caption&"' size='60' />","75%","")
		S=S&BBS.Row("<b>所在版块：</b>",Temp,"75%","")
		If BBS.Info(60)="1" Then Temp="UbbEdit()" Else Temp="HtmlEdit()"
		Temp="<script type=""text/javascript"">"&Temp&"</script>"
		S=S&BBS.Row("<b>公告内容：</b>",Temp,"75%","")
		S=S&"<div style="" padding:5px;BACKGROUND: "&BBS.SkinsPIC(2)&";"" align=""center""><input type='button' class='button' value='提 交' id='sayb' onclick='checkform(16200)' /> <input type='reset' value=' 重 写 ' onclick='Goreset()' class='button' /></div></form>"
		BBS.ShowTable Title,S
		AdminPlacard()
	Else
		TmpBoardID=BBS.Checknum(request("BoardID"))
		IF len(Content)>16200 or Len(Caption)>250 Then BBS.GoToErr(18)
		If BBS.Fun.CheckIsEmpty(Content) Then BBS.GoToErr(50)
		If BBS.Info(60)="1" Then Content=BBS.Fun.Replacehtml(Content)
		Temp=BBS.Fun.UbbString(Content)
		If ID<>0 Then
			Title="编辑公告"
			BBS.Execute("update [Placard] Set Caption='"&Caption&"',Content='"&Content&"',BoardID="&TmpBoardID&",UbbString='"&Temp&"' where ID="&ID)
		Else
			BBS.execute("insert into[Placard](Caption,Content,AddTime,Name,BoardID,UbbString)values('"&Caption&"','"&Content&"','"&BBS.NowBbsTime&"','"&BBS.MyName&"',"&TmpBoardID&",'"&Temp&"')")
			Title="发布公告"
		End If
		Content="<li>"&Title&"成功!</li><li><a href='Index.asp'>返回首页</a></li>"
		IF TmpBoardID>0 Then Content=Content&"<li><a href=board.asp?BoardID="&BBS.BoardID&">"&BBS.BoardName&"</a>"
		Content="<div style='margin:15px;line-height: 150%'>"&Content&"</div>"
		BBS.Cache.clean("Placard")
		BBS.NetLog Title
		BBS.ShowTable Title,Content
	End If
End Sub

Sub AdminPlacard()
	Dim P,arr_Rs,i,Temp,S,PInfo,Sqlwhere,Title
	If BBS.MyAdmin=7 Then
		Sqlwhere="Name='"&BBS.MyName&"'"
		Title="我发表的公告"
	Else
		Title="公告列表"
	End If
	Set P = New Cls_PageView
	P.strTableName = "[Placard]"
	P.strPageUrl = "?Action="&Action
	P.strFieldsList = "ID,Caption,BoardID,Name,AddTime,hits"
	P.strPrimaryKey ="ID"
	p.strCondiction=Sqlwhere
	P.strOrderList = "BoardID,ID desc"
	P.intPageSize = Request.QueryString("page")
	P.strCookiesName = "Placard_List"
	P.strPageVar = "page"
	P.InitClass
	Arr_Rs = P.arrRecordInfo
	PInfo = P.strPageInfo
	Set P = nothing
	S="<div style='text-align:center;border-bottom:1px solid "&BBS.SkinsPIC(0)&";padding:3px;font-weight:bold;BACKGROUND: "&BBS.SkinsPIC(2)&";'><div style=""float:left;width:32%"">公告标题</div><div style=""float:left;width:18%"">所在版块</div><div style=""float:left;width:12%"">作者</div><div style=""float:left;width:20%"">时间</div><div style=""float:left;"">管理</div><div style=""clear: both;""></div></div>"
	If IsArray(Arr_Rs) Then
		For i = 0 to UBound(Arr_Rs, 2)
		S=S&"<div style='overflow :hidden;text-align:center;border-bottom:1px solid "&BBS.SkinsPIC(0)&";height:22px;line-height:22px'><div style=""float:left;width:32%;text-align:left;""><a href='#this' onclick=""openwin('preview.asp?Action=placard&ID="&Arr_Rs(0,i)&"',500,400,'yes')"" >"&Arr_Rs(1,i)&"</a></div><div style=""float:left;width:18%;border-left:1px solid "&BBS.SkinsPIC(0)&";"">"&Replace(BBS.GetBoardName(Arr_Rs(2,i)),"所有版块","论坛首页")&"</div><div style=""float:left;width:12%;border-left:1px solid "&BBS.SkinsPIC(0)&";""><a href='UserInfo.asp?Name="&Arr_Rs(3,i)&"'>"&Arr_Rs(3,i)&"</a></div><div style=""float:left;width:20%;border-left:1px solid "&BBS.SkinsPIC(0)&";"">"&Arr_Rs(4,i)&"</div><div style=""float:left;border-left:1px solid "&BBS.SkinsPIC(0)&";padding-left:5px;height:22px;""><a href='?Action=Edit&ID="&Arr_Rs(0,i)&"&BoardID="&Arr_Rs(2,i)&"'><img src='Images/icon/edit.gif' alt='' border='0'>修改</a> <a href='#this' onclick=""if(confirm('删除这条公告！\n\n您确定要删除吗？'))window.location.href='?action=del&ID="&Arr_Rs(0,i)&"'"" ><img src='Images/icon/del.gif' alt='' border='0'>删除</a></div><div style=""clear: both;""></div></div>"
		Next
		S=S&"<div style=""BACKGROUND: "&BBS.SkinsPIC(2)&";"">"&PInfo&"</div>"
	End If
	BBS.ShowTable Title,S
End Sub
%>
