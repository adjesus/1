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
	BBS.Head"preview.asp?Action=vote","","ͶƱ����"
	If Not BBS.FoundUser Then Response.Write"�㻹û�е�½�����ܲ鿴ͶƱ��ϸ��Ϣ��":Response.End
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
			VoteShow=VoteShow&BBS.Row(i&". "&BBS.Fun.HtmlCode(Vote(i)),"<img border=0 height=8 width='"&VotePicW&"%' src='Images/hr"&ii&".gif' /> <b>"&VoteNum(i)&"</b> Ʊ","40%","22px")
		next
	End if
	Content=VoteShow&"<div style=""padding:5px;BACKGROUND: "&BBS.SkinsPIC(2)&";"" align=""center"">��Ͷ�ˣ�"&AllvoteNum&"Ʊ&nbsp;��ֹʱ�䣺"&Rs("OutTime")&" </div>"
	Rs.Close
	Response.Write"</head><body>"
	BBS.ShowTable "ͶƱѡ��",Content
	
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
					Temp=Temp&"ͶƱ��"&VoteOpt(ii)&"�"&BBS.Fun.HtmlCode(Vote(int(VoteOpt(ii))))&"<br>"
				End if
			Next
			Content=Content&BBS.Row("&nbsp;"&Arr_Rs(1,i),Temp,"70%","22px")
		Next
		BBS.ShowTable"ͶƱ�û�",Content
	End If
End Sub
Sub HeadPic()
	Dim Content,Temp,I,tr_I
	BBS.Head"preview.asp?Action=headpic","","ͷ��ѡ����"
	Response.Write"<script language=""JavaScript"" type=""text/javascript"">function instrPic(ID){self.opener.document.getElementById(""pic"").src=""pic/headpic/""+ID+"".gif"";self.opener.document.getElementById(""picurl"").value=""pic/headpic/""+ID+"".gif"";window.close();self.opener.document.getElementById(""picw"").value='"&BBS.info(54)&"';self.opener.document.getElementById(""pich"").value='"&BBS.info(55)&"'}</script>"
	Response.Write"</haed><body>"
	For I=1 To Int(BBS.Info(53))
		tr_i=tr_i+1
		Temp=Temp &"<td style='cursor:pointer' title='���ѡ�� "& I &" ��ͷ��' onclick='instrPic("&I&")'><img Src='Pic/HeadPic/"& i &".Gif'></td>"
		If tr_i=5 Then Temp=temp &"</tr><tr>":Tr_i=0
	Next
	Content="<table width='100%' barder=1>"& Temp &"</table>"
	BBS.ShowTable "��̳�Դ���ͷ�� ��"& BBS.Info(53) &"��",Content
End Sub

Sub Placard()
	Dim Rs,Caption,Content,IUBB,S
	BBS.Head"preview.asp?Action=placard","","��̳����"
	Response.Write"</head><body>"
	Set Rs=BBS.execute("select Caption,Content,AddTime,Name,hits,ubbString from [Placard] where Id="&ID&"")
	If Rs.eof then
		Caption="������Ϣ"
		Content="û�й������ݡ�"
	Else
		Set IUBB=New Cls_IUBB
		IUBB.UbbString=Rs("ubbString")
		Caption=BBS.Fun.HtmlCode(Rs("Caption"))
		S="<div style=""min-height:180px;text-indent: 24px;font-size:9pt;line-height:normal;margin-top:10px;word-wrap : break-word ;word-break : break-all ;"" onload=""this.style.overflowX='auto';"">"
		If BBS.MSIE Then S=Replace(S,"min-","width:97%;padding-right:0px; overflow-x: hidden;")
		Content="<blockquote>"&S&IUBB.UBB(Rs("Content"),2)&"</div></blockquote><div style=""padding:5px;BACKGROUND: "&BBS.SkinsPIC(2)&";"" align=""center"">�����ˣ�"&Rs("name")&"&nbsp;|&nbsp; �����ڣ�"&Rs("AddTime")&"&nbsp;|&nbsp;�Ķ�������"&Rs("hits")&" </div>"
		Set IUBB=Nothing
		BBS.execute("Update [Placard] set Hits=Hits+1 where Id="&ID&"")
	End If
	Rs.close
	Set Rs=Nothing
	BBS.ShowTable Caption,Content
	Response.Write"<div align='center'><input type='button' class='button' onclick='window.close();' value='�رմ���'></div>"
End Sub

Sub Preview()
	Dim Caption,Content,IUBB,S
	BBS.Head"preview.asp?Action=preview","","����Ԥ��"
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
	BBS.Head"preview.asp?Action=CheckName","","����û���"
	Name=trim(Request("name"))
	Caption="���û�ע��"
	If Name="" or Name=NULL Then
		Temp= "�Բ���!<br>����д�û���!"
	Else
		If not BBS.Fun.CheckName(Name) or BBS.Fun.strLength(Name)>14 or BBS.Fun.strLength(Name)<2  Then
			Temp="�Բ���!<br>�û��� <font color=red><b>"&BBS.Fun.HtmlCode(Name)&"</b></font><br>���зǷ��ַ����ַ���������"
		Else
			If Not BBS.execute("select name from [User] where name='"&Name&"'").eof Then
				Temp="�Բ���!<br>�û��� <span style='color:#F00'><b>"&BBS.Fun.HtmlCode(Name)&"</span></b> �ѱ���ע����!"
			Else
				If instr(lcase(BBS.Info(52)),lcase(Name))>0 Then
					Can=true
				End If
				If Can Then
					Temp="�Ƿ��û������������ַ�������ע��!"
				Else
					Temp="��ϲ�㣬<span style='color:#F00'><b>"&Name&"</b></span> ����ע�ᡣ"
				End If
			End If
		End If
	End If
	Response.Write "<div style='height:94px;width:294px;border:3px double #819A5F;background-color:#FFF'><div style='height:22px;line-height:22px;background-color:#9CB685;'>&nbsp;���û�ע��</div><div align='center'><br />"&Temp&"</div></div>"
End Sub%>
