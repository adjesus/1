<!--#include file="Inc.asp"-->
<!--#include file="inc/ubb_Cls.asp"-->
<%
Dim Action,ID,UserName,BbsID
BBS.ShowHead=false
Action=Lcase(Request.querystring("Action"))
If len(Action)>10 Then BBS.GoToErr(1)
ID=BBS.Checknum(request.querystring("ID"))
Select Case Action
Case"vote"
	Vote
Case"placard"
	Placard()
Case"preview"
	Preview()
Case"checkname"
	CheckUserName()
Case"headpic"
	HeadPic()
End Select
Response.Write"</body></html>"
Set BBS =Nothing

Sub Vote()
	Dim Rs,Arr_Rs,i,Temp,Content,Vote,VoteNum,AllvoteNum,VotePicW,ii,VoteShow,VoteType,voteopt
	BBS.Head"preview.asp?Action=vote","","投票详情"
	If Not BBS.FoundUser Then Response.Write"你还没有登陆，不能查看投票详细信息。":Response.End
	Set Rs=BBS.Execute("Select TopicID,Vote,VoteNum,VoteType,OutTime From [TopicVote] where TopicID="&ID&"")
	If Not Rs.Eof then
		VoteType=Rs("VoteType")
		Vote=Split(Rs("Vote"),"|")
		VoteNum=split(Rs("VoteNum"),"|")
		For i=1 to ubound(Vote)
			AllvoteNum=Int(AllvoteNum+VoteNum(i))
		Next
		IF AllVoteNum=0 then AllvoteNum=1
		For i=1 To ubound(Vote)
			ii=ii+1
			VotePicW=VoteNum(i)/AllvoteNum*85
			IF ii>6 Then ii=1
			VoteShow=VoteShow&BBS.Row(i&". "&BBS.Fun.HtmlCode(Vote(i)),"<img border=0 height=8 width='"&VotePicW&"%' src='Images/hr"&ii&".gif' /> <b>"&VoteNum(i)&"</b> 票","40%","22px")
		next
	End if
	Content=VoteShow&"<div style=""padding:5px;BACKGROUND: "&BBS.SkinsPIC(2)&";"" align=""center"">共投了："&AllvoteNum&"票&nbsp;截止时间："&Rs("OutTime")&" </div>"
	Rs.Close
	Response.Write"</head><body>"
	BBS.ShowTable "投票选项",Content
	
	Set Rs=BBS.execute("select VoteNum,User From[TopicVoteUser] where TopicID="&ID&"")
	Content=""
	If Not Rs.eof Then Arr_Rs=Rs.GetRows
	Rs.Close
	If IsArray(Arr_Rs) Then
		For i=0 To Ubound(Arr_Rs,2)
			VoteOpt=Split(Arr_Rs(0,i),",")
			Temp=""
			For II=0 to ubound(VoteOpt)
				If VoteOpt(ii)<>"" then
					Temp=Temp&"投票第"&VoteOpt(ii)&"项："&BBS.Fun.HtmlCode(Vote(int(VoteOpt(ii))))&"<br>"
				End if
			Next
			Content=Content&BBS.Row("&nbsp;"&Arr_Rs(1,i),Temp,"70%","22px")
		Next
		BBS.ShowTable"投票用户",Content
	End If
End Sub
Sub HeadPic()
	Dim Content,Temp,I,tr_I
	BBS.Head"preview.asp?Action=headpic","","头像选择器"
	Response.Write"<script language=""JavaScript"" type=""text/javascript"">function instrPic(ID){self.opener.document.getElementById(""pic"").src=""pic/headpic/""+ID+"".gif"";self.opener.document.getElementById(""picurl"").value=""pic/headpic/""+ID+"".gif"";window.close();self.opener.document.getElementById(""picw"").value='"&BBS.info(54)&"';self.opener.document.getElementById(""pich"").value='"&BBS.info(55)&"'}</script>"
	Response.Write"</haed><body>"
	For I=1 To Int(BBS.Info(53))
		tr_i=tr_i+1
		Temp=Temp &"<td style='cursor:pointer' title='点击选择 "& I &" 号头像' onclick='instrPic("&I&")'><img Src='Pic/HeadPic/"& i &".Gif'></td>"
		If tr_i=5 Then Temp=temp &"</tr><tr>":Tr_i=0
	Next
	Content="<table width='100%' barder=1>"& Temp &"</table>"
	BBS.ShowTable "论坛自带的头像 共"& BBS.Info(53) &"个",Content
End Sub

Sub Placard()
	Dim Rs,Caption,Content,IUBB,S
	BBS.Head"preview.asp?Action=placard","","论坛公告"
	Response.Write"</head><body>"
	Set Rs=BBS.execute("select Caption,Content,AddTime,Name,hits,ubbString from [Placard] where Id="&ID&"")
	If Rs.eof then
		Caption="错误信息"
		Content="没有公告内容。"
	Else
		Set IUBB=New Cls_IUBB
		IUBB.UbbString=Rs("ubbString")
		Caption=BBS.Fun.HtmlCode(Rs("Caption"))
		S="<div style=""min-height:180px;text-indent: 24px;font-size:9pt;line-height:normal;margin-top:10px;word-wrap : break-word ;word-break : break-all ;"" onload=""this.style.overflowX='auto';"">"
		If BBS.MSIE Then S=Replace(S,"min-","width:97%;padding-right:0px; overflow-x: hidden;")
		Content="<blockquote>"&S&IUBB.UBB(Rs("Content"),2)&"</div></blockquote><div style=""padding:5px;BACKGROUND: "&BBS.SkinsPIC(2)&";"" align=""center"">发布人："&Rs("name")&"&nbsp;|&nbsp; 发表于："&Rs("AddTime")&"&nbsp;|&nbsp;阅读次数："&Rs("hits")&" </div>"
		Set IUBB=Nothing
		BBS.execute("Update [Placard] set Hits=Hits+1 where Id="&ID&"")
	End If
	Rs.close
	Set Rs=Nothing
	BBS.ShowTable Caption,Content
	Response.Write"<div align='center'><input type='button' class='button' onclick='window.close();' value='关闭窗口'></div>"
End Sub

Sub Preview()
	Dim Caption,Content,IUBB,S
	BBS.Head"preview.asp?Action=preview","","帖子预览"
	Response.Write"</head><body>"
	Caption=BBS.Fun.HtmlCode(trim(request.form("PCaption")))
	Content=BBS.CheckEspecial(request.form("PContent"))
	S="<div style=""min-height:180px;text-indent: 24px;font-size:9pt;line-height:normal;margin-top:10px;word-wrap : break-word ;word-break : break-all ;"" onload=""this.style.overflowX='auto';"">"
	If BBS.MSIE Then S=Replace(S,"min-","width:97%;padding-right:0px; overflow-x: hidden;")
	Set IUBB=New Cls_IUBB
	IUBB.UbbString=BBS.Fun.UbbString(Content)
	Content=S&"<blockquote>"&S&IUBB.UBB(Content,1)&"</blockquote></div>"
	Set IUBB=Nothing
	BBS.ShowTable Caption,Content
End Sub

Sub CheckUserName()
	Dim Caption,Content,Temp,Name,can,I
	BBS.Head"preview.asp?Action=CheckName","","检测用户名"
	Name=trim(Request("name"))
	Caption="新用户注册"
	If Name="" or Name=NULL Then
		Temp= "对不起!<br>请填写用户名!"
	Else
		If not BBS.Fun.CheckName(Name) or BBS.Fun.strLength(Name)>14 or BBS.Fun.strLength(Name)<2  Then
			Temp="对不起!<br>用户名 <font color=red><b>"&BBS.Fun.HtmlCode(Name)&"</b></font><br>含有非法字符或字符过多或过少"
		Else
			If Not BBS.execute("select name from [User] where name='"&Name&"'").eof Then
				Temp="对不起!<br>用户名 <span style='color:#F00'><b>"&BBS.Fun.HtmlCode(Name)&"</span></b> 已被人注册了!"
			Else
				If instr(lcase(BBS.Info(52)),lcase(Name))>0 Then
					Can=true
				End If
				If Can Then
					Temp="非法用户名或含有屏蔽字符，不能注册!"
				Else
					Temp="恭喜你，<span style='color:#F00'><b>"&Name&"</b></span> 可以注册。"
				End If
			End If
		End If
	End If
	Response.Write "<div style='height:94px;width:294px;border:3px double #819A5F;background-color:#FFF'><div style='height:22px;line-height:22px;background-color:#9CB685;'>&nbsp;新用户注册</div><div align='center'><br />"&Temp&"</div></div>"
End Sub%>
