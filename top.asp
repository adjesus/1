<!--#include file="Inc.asp"-->
<%
Dim Action,num,Border,trHeight,Bordercolor,Bgcolor,BBSURL,Show,Face,tg,slen
Num=request.querystring("num")
Border=request.querystring("bo")
bordercolor=Left(request.querystring("boc"),9)
bgcolor=Left(request.querystring("bgc"),9)
trHeight=Left(request.querystring("h"),2)
Face=Left(request.querystring("face"),2)
tg=left(request.querystring("tg"),2)
slen=Request.QueryString("len")
If Num="" or Not BBS.Fun.isInteger(Num) Then Num="10"
If Int(Num)>50 then Num="50"
If slen="" Or Not BBS.Fun.isInteger(slen) Then slen=25
If Border="" or not BBS.Fun.isInteger(Border) Then Border="0"
If bgcolor<>"" Then bgcolor="bgcolor='#"&bgcolor&"' "
If bordercolor<>"" Then bordercolor="bordercolor='#"&bordercolor&"'"
If trHeight="" or not BBS.Fun.isInteger(trHeight) Then trHeight="18"

Action=Request.querystring("action")
If len(Action)>10 Then Response.Write"请检查调用语句":Response.End
Select Case Action
Case"topic"
Topic
Case"user"
User
Case"placard"
Placard
Case"board"
Board
Case"info"
Info
Case"login"
Login
End Select
Set BBS =Nothing

Sub User
	Dim Sql,Rs,Sqlwhere,Sqlorder,Flag,uNum
	Flag=Int(request.querystring("Flag"))
	uNum=0
	Select Case Flag
	Case 1'发帖冠军
		Sqlorder="EssayNum DESC"
	Case 2'金钱
		Sqlorder="Coin DESC"
	Case 3'积分王
		Sqlorder="Mark DESC"
	Case 4'游戏币
		Sqlorder="GameCoin DESC"
	Case Else
		Sqlorder="Regtime DESC"
	End Select
	Sql="SELECT TOP "&Num&" Name,ID FROM [User] WHERE isdel=0 ORDER BY "&Sqlorder&""
	Set Rs=BBS.Execute(Sql)
	Do while Not Rs.eof
	uNum=uNum+1
	Show=Show&"<tr height="&trHeight&"><td>"&face&" <a target='_blank' href='"&BBSURL&"/userinfo.asp?name="&Rs(0)&"'>"&BBS.Fun.StrLeft(Rs(0),Int(slen))&"</a></td></tr>"
	If uNum>=Int(Num) Then Exit Do
	Rs.movenext
	Loop
	Rs.Close
	Set Rs=Nothing
End Sub

Sub Login
Show="<form method='POST' action='"&BBSURL&"/login.asp?action=login' style='margin:0'"
If tg="1" then Show=Show&" target='_blank' "
Show=Show&"><tr height='"&trHeight&"'><td align='center'>用户：<input size='8' name='name' class='text' /></td></tr><tr><td align='center'> 密码：<input type='password' size='8' name='password' class='text'  /></td></tr>"
If request.querystring("CK")="1" Then Show=Show&"<tr height='"&trHeight&"'><td align='center'>保存：<SELECT size=1 name='cookies'><OPTION value=0 selected>不保存</OPTION><OPTION value=1>保存一天</OPTION><OPTION value=30>保存一月</OPTION><PTION value=365>保存一年</OPTION></SELECT></td></tr>"
If request.querystring("HI")="1" Then Show=Show&"<tr height='"&trHeight&"'><td align='center'>方式：<SELECT size=1 name='hidden'><OPTION value='1' selected>正常登陆</OPTION><OPTION value=2>隐身登陆</OPTION></SELECT></td></tr>"
If BBS.Info(14)="1" Then Show=Show&"<tr height='"&trHeight&"'><td>"&BBS.GetiCode&"</td></tr>"
Show=Show&"<tr height='"&trHeight&"'><td align='center'><input type='submit' class='button' value='登 陆' /> <input type='button' class='button' onClick=window.location.href='"&BBSUrl&"/Register.asp' value='注 册' /></td></tr></form>"
End Sub

Sub Topic
	Dim Sql,Rs,ARs,Sqlwhere,Sqlorder,i,Order,BoardID,TopicType,UserName,ShowTime,DayBound
	Dim Border,Height
	BoardID=request.querystring("boardid")
	TopicType=request.querystring("type")
	Order=request.querystring("order")
	UserName=request.querystring("user")
	ShowTime=request.querystring("time")
	DayBound=request.querystring("day")
	If Order="1" Then
		Sqlorder="AddTime DESC"
	ElseIf Order="2" Then
		Sqlorder="ReplyNum DESC"
	ElseIf Order="3" Then
		Sqlorder="Hits DESC"
	Else
		Sqlorder="LastTime DESC"
	End If
	Sqlwhere="isdel=0"
	If DayBound<>"" And BBS.Fun.isInteger(DayBound) then
		Sqlwhere=Sqlwhere&" And DATEDIFF('d',[AddTime],'"&BBS.NowBbsTime&"')<"&DayBound
	End If
	If BoardID<>"" And BBS.Fun.isInteger(BoardID) and BoardID<>"0" Then 
		Sqlwhere=Sqlwhere&" and boardid="&BoardID&""
	End If
	Select Case Int(TopicType)
	Case 1
	Sqlwhere=Sqlwhere&" AND topType>0"
	Case 2
	Sqlwhere=Sqlwhere&" AND isGood=1"
	Case 3
	Sqlwhere=Sqlwhere&" AND isVote=1"
	End Select 
	Sql="SELECT Top "&Num&" TopicID,Name,Face,Caption,BoardID,LastTime,SqlTableID FROM [Topic] WHERE "&Sqlwhere&" ORDER BY "&Sqlorder&""
	Set Rs=BBS.Execute(Sql)
	If Not Rs.Eof Then
		ARs=Rs.GetRows(Num)
		Rs.Close
		Set Rs=Nothing
		For I=0 To Ubound(ARs,2)
			Show=Show&"<tr height='"&trHeight&"'>"
			If Face="1" Then Show=Show&"<td width='20'><img src='"&BBSURL&"/pic/face/"&ARs(2,i)&".gIf' alt='' /></td>"
			Show=Show&"<td width=*>"
			If Face<>"1" Then
			If Face="0" Then
				Show=Show&i+1&"."
			Else
				Show=Show&Face
			End If
			End If
			Show=Show&" <a"
			If tg="1" then Show=Show&" target='_blank' "
			Show=Show&" href='"&BBSURL&"/topic.asp?id="&ARs(0,i)&"&BoardID="&ARs(4,i)&"&TB="&ARs(6,i)&"'>"&BBS.Fun.StrLeft(ARs(3,i),Int(slen))&"</a></td>"
			If UserName="1" Then Show=Show&"<td width='110' align='center'>"&ARs(1,i)&"</td>"
			If ShowTime="1" Then Show=Show&"<td width='120' align='center'>"&ARs(5,i)&"</td>"
			Show=Show&"</tr>"
		Next
	Else
		Show="<tr height="&trHeight&"><td>没有内容</td></tr>"
	End If
End Sub

Sub Board
	Dim I,II,po
	If Not IsArray(BBS.Board_Rs) Then BBS.GetBoardCache()
	If IsArray(BBS.Board_Rs) Then
		For i=0 To Ubound(BBS.Board_Rs,2)
			po=""
			If BBS.Board_Rs(0,i)<>0  Then
				For II=1 To BBS.Board_Rs(0,i)
				Po=Po&"&nbsp;"
				Next
			End If
			Show=Show&"<tr height="&trHeight&"><td>"&Po&" ├ <a "
			If tg="1" Then Show=Show&"target='_blank' "
			Show=Show&"href='"&BBSURL&"/board.asp?boardid="&BBS.Board_Rs(1,i)&"'>"&BBS.Board_Rs(3,i)&"</a></td></tr>"
		Next
	End If
End Sub

'公告
Sub Placard()
	Dim Temp,Rs,Arr_Rs,i,BoardID,ShowTime
	BoardID=request.querystring("boardid")
	ShowTime=request.querystring("time")
	If BBS.Cache.valid("Placard") then
		Arr_Rs=BBS.Cache.Value("Placard")
	Else
		Set Rs=BBS.Execute("Select Id,Caption,AddTime,BoardID From [Placard] order by Id desc")
		If Rs.Eof Or Rs.Bof Then
			Show="<tr height="&trHeight&"><td>没有公告</td></tr>"
			Exit sub
		Else
			Arr_Rs=Rs.GetRows(-1)
			Rs.Close
			Set rs=nothing
			BBS.Cache.add "Placard",Arr_Rs,dateadd("n",5000,now)
		End if
	End if
	For i=0 To Ubound(Arr_Rs,2)
	If i>int(num) then exit for
		Temp="<tr height="&trHeight&"><td>"&face&" <a "
		If tg="1" Then Temp=Temp&"target='_blank' "
		Temp=Temp&"href='"&BBSURL&"/preview.asp?Action=placard&id="&Arr_Rs(0,i)&"'>"&BBS.Fun.StrLeft(BBS.Fun.HtmlCode(Arr_Rs(1,i)),Int(slen))&"</a></td>"
		If ShowTime="1" Then Temp=Temp&"<td width='80' align='center' style='font-size:12px'>"&Arr_Rs(2,i)&"</td>"
		Temp=Temp&"</tr>"
	If BoardID="" Then
		Show=Show&Temp
	Else
		If Int(Arr_Rs(3,i))=int(boardID) Then
		Show=Show&Temp
		End If
	End If
	Next
	If Show="" Then Show="<tr height="&trHeight&"><td>没有公告</td></tr>"
End Sub

Sub Info
Dim flag
flag=request.querystring("flag")
BBS.Getonline'读取在线
If instr(flag,"|1|")=0 then Show=Show&"<tr height="&trHeight&"><td>"&face&" 论坛帖数：<font color=red>"&BBS.InfoUpdate(0)&"</a></td></tr>"
If instr(flag,"|2|")=0 then Show=Show&"<tr height="&trHeight&"><td>"&face&" 主题帖数：<font color=red>"&BBS.InfoUpdate(1)&"</a></td></tr>"
If instr(flag,"|3|")=0 then Show=Show&"<tr height="&trHeight&"><td>"&face&" 今日帖数：<font color=red>"&BBS.InfoUpdate(2)&"</a></td></tr>"
If instr(flag,"|4|")=0 then Show=Show&"<tr height="&trHeight&"><td>"&face&" 昨日帖数：<font color=red>"&BBS.InfoUpdate(3)&"</a></td></tr>"
If instr(flag,"|5|")=0 then Show=Show&"<tr height="&trHeight&"><td>"&face&" 最高日帖：<font color=red>"&BBS.InfoUpdate(4)&"</a></td></tr>"
If instr(flag,"|6|")=0 then Show=Show&"<tr height="&trHeight&"><td>"&face&" 注册人数：<font color=red>"&BBS.InfoUpdate(5)&"</a></td></tr>"
If instr(flag,"|7|")=0 then Show=Show&"<tr height="&trHeight&"><td>"&face&" 最新会员：<a target='_blank' href='"&BBSURL&"/Userinfo.asp?name="&BBS.InfoUpdate(6)&"'><font color=red>"&BBS.InfoUpdate(6)&"</a></td></tr>"
If instr(flag,"|8|")=0 then Show=Show&"<tr height="&trHeight&"><td>"&face&" 论坛在线：<font color=red>"&BBS.AllonlineNum&"</a></td></tr>"
If instr(flag,"|9|")=0 then Show=Show&"<tr height="&trHeight&"><td>"&face&" 在线会员：<font color=red>"&BBS.useronlineNum&"</a></td></tr>"
If instr(flag,"|10|")=0 then Show=Show&"<tr height="&trHeight&"><td>"&face&" 在线游客：<font color=red>"&BBS.AllonlineNum-BBS.useronlineNum&"</a></td></tr>"
If instr(flag,"|11|")=0 then Show=Show&"<tr height="&trHeight&"><td>"&face&" 最高在线：<font color=red>"&BBS.InfoUpdate(7)&"</a></td></tr>"
If instr(flag,"|12|")=0 then Show=Show&"<tr height="&trHeight&"><td>"&face&" 建站时间：<font color=red>"&BBS.Info(5)&"</a></td></tr>"
End Sub
%>
document.write("<table width='100%' border='<%=border%>' <%=bgcolor&bordercolor%> cellpadding='0' cellspacing='0' style='border-collapse:collapse'><%=Show%></table>");