<!--#include file="inc.asp"--><%
Dim Action,ID,Page,Temp
Dim Caption,SubmitUrl
Dim Title,Content
If Not BBS.Founduser Then BBS.GoToerr(31)
BBS.CheckBoard()
ID=BBS.CheckNum(request.querystring("ID"))
Page=BBS.CheckNum(request.querystring("page"))
Action=lcase(request.querystring("action"))
If Len(Action)>10 Then BBS.GoToerr(1)
If Session(CacheName & "MyGradeInfo")(10)="1" Then 
	Temp=" <select name='font_s'><option value=''>��ʽ</option><option value='B'>����</option><option value='I'>б��</option><option value='U'>����</option><option value='no'>Ĭ��</option></select><select name='font_c'><option value=''>��ɫ</option><option style='COLOR:#000;BACKGROUND-COLOR:#000' value='#000'></option><option style='COLOR:#000;BACKGROUND-COLOR:#F00' value='#F00'></option><option style='COLOR:#000;BACKGROUND-COLOR:#00F' value='#00F'></option><option style='COLOR:#000;BACKGROUND-COLOR:#0F0' value='#0F0'></option><option style='COLOR:#000;BACKGROUND-COLOR:#008000' value='#008000'></option><option style='COLOR:#000;BACKGROUND-COLOR:#FA0' value='#FA0'></option><option style='COLOR:#000;BACKGROUND-COLOR:#F0F' value='#F0F'></option><option style='COLOR:#000;BACKGROUND-COLOR:#0FF' value='#0FF'></option><option style='COLOR:#000;BACKGROUND-COLOR:#888' value='#888'></option><option style='COLOR:#000;BACKGROUND-COLOR:#800000' value='#800000'></option><option style='COLOR:#000;BACKGROUND-COLOR:#800080' value='#800080'></option><option style='COLOR:#000;BACKGROUND-COLOR:#008080' value='#008080'></option><option style='COLOR:#000;BACKGROUND-COLOR:#000080' value='#000080'></option><option style='COLOR:#000;BACKGROUND-COLOR:#808000' value='#808000'></option><option value='no'>Ĭ��</option></select>"
End If
Title=BBS.Row("<b>�������⣺</b><select name='Title' onChange='DoTitle(this.options[this.selectedIndex].value)' style='font-size: 9pt'><option selected value=''>����</option><option value='[ԭ��]'>[ԭ��]</option><option value='[ת��]'>[ת��]</option><option value='[��ˮ]'>[��ˮ]</option><option value='[����]'>[����]</option><option value='[����]'>[����]</option><option value='[�Ƽ�]'>[�Ƽ�]</option><option value='[����]'>[����]</option><option value='[ע��]'>[ע��]</option><option value='[��ͼ]'>[��ͼ]</option><option value='[����]'>[����]</option><option value='[����]'>[����]</option><option value='[����]'>[����]</option></select>","<input id='caption' name='caption' type='text' style='width:60%' maxlength='200' />"&Temp,"75%","")

Select Case Action
	Case"vote"
		Vote()
	Case"reply"
		Reply()
	Case"edit"
		Edit()
	Case Else
		BBS.Stats="��������"
		Submiturl="postsave.asp?boardid="&BBS.boardid
End Select
BBS.Head "post.asp?boardid="&BBS.boardid,BBS.BoardName,BBS.Stats
ShowMain()
BBS.Footer()
Set BBS =Nothing

Sub Vote()
	Dim i
	If Session(CacheName & "MyGradeInfo")(12)="0" Then
		Temp="<div style=""padding:4px"">�Բ�����Ŀǰ����̳�ȼ�û�з���ͶƱ�����Ȩ�ޡ�</div>"
	Else
		Temp="��ѡ��ͶƱ��Ŀ����<select name='votenum' id='votenum' onchange='SetNum(this)' />"
		For i = 2 to int(BBS.Info(63))
			Temp=Temp&"<option value='"&I&"'>"&I&"</option>"
		Next
		Temp=Temp&"</select>�����ѡ<input type='checkbox' name='votetype' value='2' /> ����ʱ�䣺<select name='outtime'><option value='1'>һ��</option><option value='3'>����</option><option value='7'>һ��</option><option value='15'>�����</option><option value='31'>һ����</option><option value='93'>������</option><option value='365'>һ��</option><option value='10000' selected>������</option></select><hr size=1 width='98%' /><div id='optionid'><div>ѡ��1��<input type='text' name='Votes1' style='width:80%' /></div><div>ѡ��2��<input type='text' name='votes2' style='width:80%' /><INPUT TYPE='hidden' name='autovalue' value='2' /></div></div>"
	End If	
	Title=Title&BBS.Row("<b>ͶƱѡ�</b>",Temp,"75%","")
	BBS.Stats="������ͶƱ"
	SubmitUrl="postsave.asp?boardid="&BBS.boardid
End Sub

Sub Reply()
	Dim Rs,BbsID
	if ID=0 Then BBS.GoToErr(1)
	BBS.Stats="�ظ�����"
	Set Rs=BBS.Execute("Select Caption,SqlTableID,IsLock,IsDel From [Topic] where TopicID="&ID&" And IsDel=0")
	If Rs.Eof Then
		BBS.GoToErr(21)
	ElseIf Rs(2)=1 Then
		BBS.GoToErr(22)
	Else
		Title=BBS.Row("<input type=hidden name='caption' id='caption' value='Re:"&Rs(0)&"' />�ظ����⣺",Rs(0),"75%","22px")
		BBS.TB=Rs(1)
	End If
	Rs.close
	Set Rs=Nothing
	Submiturl="postsave.asp?Action=Reply&boardid="&BBS.boardid&"&TB="&BBS.TB&"&ID="&ID&"&page="&page

	BbsID=BBS.CheckNum(Request.querystring("BbsID"))
	If BbsID>0 Then
		Set Rs=BBS.Execute("select top 1 B.ReplyTopicID,B.TopicID,B.Name,B.AddTime,B.Content,B.boardid,U.IsShow from [Bbs"&BBS.TB&"] As B inner join [User] As U on B.Name=U.Name where B.BbsID="&BbsID&" And B.IsDel=0")
			If Not Rs.Eof Then
				If Rs(1)<>ID And Rs(0)<> ID Then BBS.GoToErr(1)
				If Rs(6)=1 Then
					Content="<div class=""quote"">���� "&RS(2)&" �ķ�������:<br><font color=""#F00"">�������ݲ�������</font><div><br>"
				Else
				If BBS.Info(60)="0" Then
				Content="<div class=""quote"">���������� <b>"&RS(2)&"</b></font> ��(<i>"&Rs(3)&"</i>)�ķ���<br>"&QuoteCode(Rs(4))&"</div><br><br>"
				Else
				Content="[quote]���������� [B]"&RS(2)&"[/B] ��<br>"&QuoteCode(Rs(4))&"<br>[/quote]<br>"			
				End If
				End If
			End if
			Rs.close
		Set Rs=Nothing
	End If
End Sub

Sub Edit()
	Dim Rs,BbsID,TopicIsLock,TopicRs,IsTop
	BbsID=BBS.CheckNum(request.querystring("BbsID"))
	IF BbsID=0 Or ID=0 Then BBS.GoToErr(1)
	
	Set Rs=BBS.Execute("Select boardid,TopType,SqlTableID,IsLock From [Topic] where IsDel<>1 And TopicID="&ID)
	If Rs.Eof Then
		BBS.GoToErr(58)
	Else
		TopicRs=Rs.GetRows(-1)
	End If
	Rs.Close
	Set Rs=BBS.Execute("select boardid,Name,AddTime,TopicID,Caption,Content,IsDel From [Bbs"&TopicRs(2,0)&"] where IsDel<>1 And BbsID="&BbsID&"")
	If Rs.eof  Then
		BBS.GoToErr(58)
	Else
	If lcase(BBS.MyName)=lcase(rs("name")) Then 
			If TopicRs(3,0)=1 And BBS.MyAdmin<>9 Then BBS.GoToErr(22)
			If Session(CacheName & "MyGradeInfo")(22)="0" Then
				If BBS.Info(12)<>"0" And DateDiff("s",Rs("AddTime")+BBS.Info(12)/1440,BBS.NowBbsTime)>0 Then BBS.GoToErr(34)
			End If
	Else
		If Session(CacheName & "MyGradeInfo")(24)="0" Then BBS.GoToErr(33)
		If TopicRs(1,0)=5 or TopicRs(1,0)=4 Then'������ܶ�������
			If TopicRs(0,0)<>BBS.boardid Then'������Ǳ��棬������Ȩ
				If BBS.MyAdmin=7 Then BBS.GoToErr(51)
			End If
		Else
			If BBS.MyAdmin=7 And Not BBS.IsBoardAdmin Then BBS.GoToErr(71)
		End If
	End If
		If TopicRs(1,0)=5 or TopicRs(1,0)=4 Then
			If lcase(BBS.MyName)<>lcase(rs("name")) Then

			End If
		Else
			If TopicRs(0,0)<>BBS.boardid Then BBS.GotoErr(1)
		End If

		IF Rs("TopicID")=0 Then
			Title=BBS.Row("<input type=hidden id='caption' name='caption' value='"&Rs(4)&"' /><b>�༭�ظ�����</b>",rs(4),"75%","23px")
		Else 
			Title=replace(Title,"id='caption'","id='caption' value='"&Rs(4)&"'")
		End IF
		Content=ReplaceUBB(rs(5))
	End if
	Rs.Close
	BBS.Stats="�༭����"
	Submiturl="postsave.asp?Action=Edit&ID="&ID&"&BbsID="&BbsID&"&boardid="&BBS.boardid&"&TB="&TopicRs(2,0)&"&page="&page&""
End Sub


Function ShowMain()
	With BBS
	Dim Face,I,Temp1,S1
	Temp="<form style='margin:0;' name='preview' action='preview.asp?Action=preview' method='post' target='preview'><input type='hidden' name='pcaption' /><input type='hidden' name='pcontent' /></form>"
	Temp=Temp&"<form style='margin:0;' method=POST name='say' action='"&Submiturl&"' >"
	Temp=Temp&title
	If .Info(15)="1" Then
		Temp=Temp&.Row("<b>������֤�룺</b>",.GetiCode,"75%","")
	Else
		Temp=Temp&"<input type=hidden name='iCode' id='iCode' value='BBS' />"
	End If
	Face="<input name=face type=radio value=1 checked class=checkbox /><img src='pic/face/1.gif' border='0' align='absmiddle' atl='' />&nbsp;"
	For i=2 to 18
		Face=Face&"<input type=radio value="&i&" name='face' class=checkbox /><img border=0 align='absmiddle' src='pic/face/"&i&".gif' atl='' />&nbsp;"
		if i=9 then Face=Face&"<br />"
	Next
	Temp=Temp&.Row("<textarea id='content' name='content' style='display:none'>"&Server.HtmlEnCode(Content)&"</textarea><b>��ı��飺</b><br />������ǰ��",Face,"75%","")
	If .Info(30)="0" Then
	  Temp1="����̳��ʱ�ر��ϴ����ܡ�<br>"
	 ElseIf Session(CacheName & "MyGradeInfo")(14)="0" then
	  Temp1="��Ŀǰ����̳�ȼ���û���ϴ���Ȩ�ޣ�"
	 ElseIf .BoardString(14)="0" then
	  Temp1="��������ʱ�ر��ϴ����ܡ�"
	 ElseIf .BoardString(14)="2" And Session(CacheName & "MyInfo")(17)="0" then
	  Temp1="������ֻ����VIP��Ա���ϴ�Ȩ�ޣ�"
	Else
		Temp1="<input style=""margin-top:10px"" class=""button"" type=""button"" value=""�ϴ�����"" onclick=""javascript:document.getElementById('up').style.display='block';upf.location.replace('UploadFile.asp');this.style.display='none'""> ���ϴ��ļ����ͣ�"&Replace(.Info(34)&"|"&.Info(35),"|","��")
		Temp1=Temp1&"<div id='up' style='display:none'><iframe id='upf' name='upf' scrolling='no' frameborder='0' height='22' width='100%'></iframe></div>"
	End if
	Temp=Temp&.Row("<b>�����ϴ���</b><br />ÿ���������ϴ�<font color=blue>"&Session(CacheName & "MyGradeInfo")(15)&"</font>��(���<font color=blue>"&Session(CacheName & "MyGradeInfo")(16)&"</font>KB)",Temp1,"75%","42px")	
	Temp1="<br /><a href=""javascript:CheckLength("&Session(CacheName & "MyGradeInfo")(9)&")""> �������ƣ�<font color=red>"&Session(CacheName & "MyGradeInfo")(9)&"�ֽ�</font></a><br />HTML��ǩ��<font color=red>"
	If .Info(60)="1" Then Temp1=Temp1&"��" Else Temp1=Temp1&"��"
	Temp1=Temp1&"</font><br />UBB��ǩ�� <font color=red>��</font><br />�ϴ��ļ���<font color=red>"
	If .Info(30)="0" Then Temp1=Temp1&"��" Else Temp1=Temp1&"��"
	Temp1=Temp1&"</font><br /><b>����������</b><br />"&_
	Especial("�ظ��ɼ�","Especial('[REPLY]','[\/REPLY]')",.Info(70))&_
	Especial("��Ǯ�ɼ�","Coin()",.Info(71))&"<br />"&_
	Especial("���ֿɼ�","Mark()",.Info(72))&_
	Especial("���ڿɼ�","Showdate()",.Info(73))&"<br />"&_
	Especial("�Ա�ɼ�","Sex()",.Info(74))&_
	Especial("��½�ɼ�","Especial('[LOGIN]','[\/LOGIN]')",.Info(75))&"<br />"&_
	Especial("ָ������","Name()",.Info(76))&_
	Especial("���ѹۿ�","Buypost()",.Info(77))&"<br />"&_
	Especial("�������ʶ��ת��","Code()",.Info(68))
	Temp1=Temp1&"<ul><li>���������ع��ҷ���</li><li>��ֹ�������μ�ɫ������</li></ul>"
	If .Info(60)="1" Then S1="UbbEdit()" Else S1="HtmlEdit()"
	S1="<script type=""text/javascript"">"&S1&"</script>"
	If BBS.CC(0)="1" Then
		S1=S1&"<object width='72' height='24'><param name='wmode' value='transparent' /><param name='allowScriptAccess' value='always' /><param name='movie' value='http://union.bokecc.com/flash/plugin_"&BBS.CC(4)&".swf?userID="&BBS.CC(2)&"&type=BBS' /><embed src='http://union.bokecc.com/flash/plugin_"&BBS.CC(4)&".swf?userID="&BBS.CC(2)&"&type=BBS' type='application/x-shockwave-flash' width='72' height='24' allowScriptAccess='always'></embed></object> &nbsp;"
	End IF
	If Action="edit" And Session(CacheName & "MyGradeInfo")(25)="1" Then S1= S1 & "<input name='editchalk' type='checkbox' value='No' class=checkbox>�����±༭���"
	If Action="" then
	If (BBS.MyAdmin=7 And Not BBS.IsBoardAdmin) Then
	Else
	S1=S1&"�������ã�"
	If Session(CacheName&"MyGradeInfo")(31)="1" Then S1= S1 & "<input name='top' type='checkbox' value='1' class=checkbox>�ö� "
	If Session(CacheName&"MyGradeInfo")(32)="1" Then S1= S1 & "<input name='classtop' type='checkbox' value='1' class=checkbox>���ö� "
	If Session(CacheName&"MyGradeInfo")(33)="1" Then S1= S1 & "<input name='alltop' type='checkbox' value='1' class=checkbox>���ö� "
	If Session(CacheName&"MyGradeInfo")(34)="1" Then S1= S1 & "<input name='good' type='checkbox' value='1' class=checkbox>���� "
	If Session(CacheName&"MyGradeInfo")(35)="1" Then S1= S1 & "<input name='lock' type='checkbox' value='1' class=checkbox>����"
	End If
	End If
	Temp=Temp&.Row("<b>�������ݣ�</b>"&Temp1,S1,"75%","")
	Temp=Temp&"<div align='center' style="" padding:5px;BACKGROUND: "&.SkinsPIC(1)&";"">"
	Temp=Temp&"&nbsp;<input type='button' value='OK ����' id='sayb' onclick='checkform("&Session(CacheName & "MyGradeInfo")(9)&")' class='button' /> <input type=button value='Ԥ ��' onclick='Gopreview()' class='button' /> <input type='reset' value='NO ��д' onclick='Goreset()' class='button' />" 
	Temp=Temp&"</div></form>"
	.ShowTable .Stats,Temp
	End With
End Function

Function replaceUBB(str)
	dim re
	If Str="" Then Exit Function
	Set re=new RegExp
	re.IgnoreCase=true
	re.Global=True
	re.Pattern="(>)("&vbNewLine&")(<)"
	Str=re.Replace(Str,"$1$3")
	re.Pattern="(>)("&vbNewLine&vbNewLine&")(<)"
	Str=re.Replace(Str,"$1$3")
	re.Pattern=vbNewLine
	Str=re.Replace(Str,"<br>")	
	re.Pattern="(\[right\])(\[color=(.[^\[]*)\])(.[^\[]*)(\[\/color\])(\[\/right\])"
	str=re.Replace(str," ")
	re.Pattern="(<div style=""color:#999999;text-align:right"">�������ӱ�(.*)�༭����<\/div>)"
	str=re.Replace(str," ")
	str=Replace(Str,"  ","&nbsp;&nbsp;")
	Set re=Nothing
	replaceUBB=str
End function

Function Especial(eName,gourl,Flag)
	If flag="1" Then
		Especial="<a href=""javascript:"&Gourl&""">"&eName&"</a> <span style='color:#F00'>��</span> "
	Else
		Especial=eName&" <span style='color:#AAA'>��</span> "
	End If
End Function

Function QuoteCode(str)
		Dim re,restr
		Set re=new RegExp
		re.IgnoreCase=true
		re.Global=True
		restr="<hr>�������ݲ������� <hr>"
		re.Pattern="(\[DATE=(.[^\[]*)\])(.+?)(\[\/DATE\])"
		str=re.Replace(str,restr)
		re.Pattern="(\[SEX=*([0-1]*)\])(.+?)(\[\/SEX\])"
		str=re.Replace(str,restr)
		re.Pattern="(\[COIN=*([0-9]*)\])(.+?)(\[\/COIN\])"
		str=re.Replace(str,restr)		
		re.Pattern="(\[USERNAME=(.[^\[]*)\])(.+?)(\[\/USERNAME\])"
		str=re.Replace(str,restr)	
		re.Pattern="(\[GRADE=*([0-9]*)\])(.+?)(\[\/GRADE\])"
		str=re.Replace(str,restr)	
		re.Pattern="(\[MARK=*([0-9]*)\])(.+?)(\[\/MARK\])"		
		str=re.Replace(str,restr)
		re.Pattern="(\[BUYPOST=*([0-9]*)\])(.+?)(\[\/BUYPOST\])"
		str=re.Replace(str,restr)
		re.Pattern=vbcrlf&vbcrlf&vbcrlf&"(\[RIGHT\])(\[COLOR=(.[^\[]*)\])(.[^\[]*)(\[\/COLOR\])(\[\/RIGHT\])"
		str=re.Replace(str,"")
		re.Pattern="(\[reply\])(.+?)(\[\/reply\])"
		Str=re.Replace(str,restr)	
		QuoteCode=replaceUBB(str)
		Set re=Nothing
End Function
%>