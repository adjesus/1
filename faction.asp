<!-- #include file="Inc.asp" -->
<!-- #include file="Inc/Page_Cls.asp" -->
<%Dim ID,Rs,Page_Url
If Not BBS.Founduser Then BBS.GoToErr(10)
If Request.QueryString("page") > 1 Then
  Page_Url = "?Page="&Request.QueryString("page")
Else
  Page_Url = ""
End If
BBS.Head"faction.asp"&Page_Url,"","��̳����"
ID = BBS.CheckNum(Request.querystring("ID"))
Select Case Left(Request("Action"),10)
Case"Add"
	Add
Case"Edit"
	Edit
Case"FactionAdd"
	FactionAdd
Case"FactionOut"
	FactionOut
Case"Look"
	Look
Case"Del"
	Del
Case Else
	Main()
End Select
BBS.Footer
Set BBS =Nothing


Sub Main()
Dim intPageNow,strPageInfo,arr_Rs,i,Pages,page,Content
Content="<table width='95%' border=0 style='border-collapse: collapse;WORD-BREAK: break-all;'><tr><td align='center' height=40 colspan=5><a href='?action=Add'><img src='Images/icon/right.gif' border='0' align='absmiddle'> ��������</a>&nbsp;&nbsp;<a href=#this onclick=""if(confirm('��ȷ��Ҫ�˳��ð��ɣ�\n\n����"&BBS.Info(121)&"�����ȥ 1'))window.location.href='?action=FactionOut'""><img src='Images/icon/right.gif'  border='0' align='absmiddle'> �˳�����</a></td></tr>"&_
"<tr><td width='15%' class=FactionTit>�ɱ�</td><td width='40%' class=FactionTit>��ּ</td><td width='15%' class=FactionTit>��ʼ��</td><td width='10%' class=FactionTit>����</td><td width='20%' class=FactionTit>��������</td></tr>"
	intPageNow = Request.QueryString("page")
	Set Pages = New Cls_PageView
	Pages.strTableName = "[Faction]"
	Pages.strFieldsList = "ID,Name,Note,User,BuildDate"
	Pages.strOrderList = "ID desc"
	Pages.strPrimaryKey = "ID"
	Pages.intPageSize = 15
	Pages.intPageNow = intPageNow
	Pages.strCookiesName = "Faction"'�ͻ��˼�¼����
	Pages.Reloadtime=3'ÿ�����Ӹ���Cookies
	Pages.strPageVar = "page"
	Pages.InitClass
	Arr_Rs = Pages.arrRecordInfo
	strPageInfo = Pages.strPageInfo
	Set Pages = nothing
	If IsArray(Arr_Rs) Then
	For i = 0 to UBound(Arr_Rs, 2)
		Content=Content & "<tr><td align='center' height='25'><a href=?action=Look&ID="&Arr_Rs(0,i)&">"&Arr_Rs(1,i)&"</a></td><td>"&Arr_Rs(2,i)&"</td><td align='center' height='25'><a href=UserInfo.asp?Name="&Arr_Rs(3,i)&">"&Arr_Rs(3,i)&"</a></td><td align='center'>"
		If SESSION(CacheName & "MyInfo")(25)=Arr_Rs(1,i) then
			Content=Content & "<a href=#this onclick=""if(confirm('��ȷ��Ҫ�˳��ð��ɣ�\n\n����"&BBS.Info(121)&"�����ȥ 1'))window.location.href='?action=FactionOut&ID="&Arr_Rs(0,i)&"'"">�˳��˰�</a>"
		Else
			Content=Content & "<a href=#this onclick=""if(confirm('��ȷ��Ҫ����ð��ɣ�\n\n����"&BBS.Info(121)&"����ﵽ 3'))window.location.href='?action=FactionAdd&ID="&Arr_Rs(0,i)&"'"">����˰�</a>"
		End if
		Content=Content & "<td align='center'><a href='?action=Edit&ID="&Arr_Rs(0,i)&"'><img src='Images/icon/edit.gif' border='0' alt='' />�޸�</a> <a href=#this onclick=""if(confirm('��ȷ��Ҫ��ɢ�ð��ɣ�'))window.location.href='?action=Del&ID="&Arr_Rs(0,i)&"'""><img src='Images/icon/del.gif' border='0' alt='' />��ɢ</a></td></tr>"
	Next
	End If
	Content=Content & "</table><div style=""padding:5px;BACKGROUND: "&BBS.SkinsPIC(2)&";"">"&strPageInfo&"<br><br></div>"

	BBS.ShowTable"��̳����",Content
End Sub

Sub FactionAdd
	Dim Content,Rs
	BBS.CheckMake
	If SESSION(CacheName & "MyInfo")(25)<>"" Then
		BBS.Alert"���Ѿ�����["&SESSION(CacheName & "MyInfo")(25)&"]�ˣ������˳�["&SESSION(CacheName & "MyInfo")(25)&"]���ܼ����°�","?"
	ElseIf Int(SESSION(CacheName & "MyInfo")(6))<3 then
		BBS.Alert"����"&BBS.Info(121)&"ֵС��3��û���ʸ������ɣ�","?"
	Else
		Set Rs=BBS.Execute("select Name from [Faction] where ID="&ID)
		IF Not Rs.Eof Then
			BBS.Execute("update [user] Set Faction='"&rs(0)&"' where Name='"&BBS.MyName&"'")
			Session(CacheName & "MyInfo") = Empty
			BBS.Alert"�ɹ��ļ��� ["&Rs(0)&"] ���ɣ�","?"
		Else
			BBS.Alert"������������ɣ�","?"
		End If
		Rs.Close
	End If
End Sub

Sub FactionOut
	BBS.CheckMake
	If SESSION(CacheName & "MyInfo")(25)="" Then
		BBS.Alert"��Ŀǰ��û�м����κΰ��ɣ�","?"
	Else
		If Not BBS.Execute("select ID from [Faction] where user='"&BBS.MyName&"'").eof Then
			BBS.Alert"���������ˣ������˳����ɣ��˳�������Ҫ��ɢ���ɣ�","?"
		Else
			BBS.execute("Update [user] Set Faction='',Mark=Mark-1 where name='"&BBS.MyName&"'")
			Session(CacheName & "MyInfo") = Empty
		End If
		BBS.Alert"�˳����ɳɹ�","?"
	End If
End Sub

Sub Del
BBS.CheckMake
	Set Rs=BBS.Execute("Select Name,User From[Faction] where ID="&ID)
	If Rs.Eof Then
		BBS.Alert"������������ɣ�","?"
	ElseIf BBS.MyName<>Rs(1) Then
		BBS.Alert"�����Ǹð�İ����޷���ɢ�ð","?"
	Else
		BBS.Execute("Update [user] set Faction='' where Faction='"&rs(0)&"'")
		BBS.Execute("Delete from [Faction] where ID="&ID)
		Session(CacheName & "MyInfo") = Empty
		BBS.Alert"��ɢ���ɳɹ���","?"
	End if
	Rs.Close
End Sub

Sub Look
Dim Content
Set Rs=BBS.Execute("Select Name,FullName,Note,User,BuildDate from [Faction] where ID="&ID)
If Rs.eof Then
	BBS.Alert"�����ڴ˰��ɣ�","?"
Else
	Content="<table width='95%' border=0 style='border-collapse: collapse;WORD-BREAK: break-all;'><tr><td width='24%' align='right' height=25><b>�������ƣ�</b></td><td width='74%'>&nbsp;"&BBS.Fun.HtmlCode(rs(0))&"</td></tr>"&_
	"<tr><td align='right' height=25><b>����ȫ�ƣ�</b></td><td>&nbsp;"&BBS.Fun.HtmlCode(rs(1))&"</td></tr>"&_
	"<tr><td align='right' height=25><b>������ּ��</b></td><td>&nbsp;"&BBS.Fun.HtmlCode(rs(2))&"</td></tr>"&_
	"<tr><td align='right' height=25><b>����ʱ�䣺</b></td><td>&nbsp;"&Rs(4)&"</td></tr>"&_
	"<tr><td align='right' height=25><b>�������ƣ�</b></td><td>&nbsp;"&Rs(3)&"</td></tr>"&_
	"<tr><td align='right' height=25><b>���е��ӣ�</b></td><td>"&Desciple(Rs(0))&"</td></tr>"&_
	"<tr><td colspan=2 align='center' height=25><a href='?'>�����ء�</a></td></tr></table>"
	BBS.ShowTable"������Ϣ",Content
End If
Rs.Close
End Sub

Function Desciple(name)
	Dim dRs,I
	I=0
	Set dRs=BBS.Execute("Select Name From [user] where Faction='"&Name&"'")
	Do while not dRs.eof
	I=I+1
	Desciple=Desciple & "<a target='_blank' href='Userinfo.asp?Name="&dRs(0)&"'>"&dRs(0)&"</a>&nbsp;&nbsp;&nbsp;"
	dRs.movenext
	Loop
	dRs.close:Set dRs=NoThing
	Desciple="<table width='95%' border='0' cellpadding='0' cellspacing='0'><tr><td>&nbsp;"&I&" ��</td><td width='90%'><marquee onmouseover='this.stop()' onmouseout='this.start()' scrollAmount='3' direction='left' width='95%' height='15'>"&Desciple&"</marquee></td></tr></table>"
End Function

Sub Add
Dim Name,FullName,Note,Content
BBS.CheckMake
Name=BBS.Fun.GetStr("Name")
FullName=BBS.Fun.GetStr("FullName")
Note=BBS.Fun.GetStr("Note")
IF Name="" And FullName="" And Note="" Then
	Content="<form  method='post' style='margin:0'>"&_
	"<table width='95%' border=0 style='border-collapse: collapse;WORD-BREAK: break-all;'><tr><td colspan=2 height=30 align='center'><font color=red>�������ɵı�Ҫ������ 1.���� "&BBS.Info(121)&" ���� 20 ���ϣ� 2.��Ҫ�۳��� 10000 ��"&BBS.Info(120)&"��Ϊ���ɻ��� </font></td></tr>"&_
	"<tr><td align='right' height=25><b>�������ƣ�</b></td><td>&nbsp;<input class='text' Maxlength=10 Name='Name' size='10'>*���ֻ��6������</td></tr>"&_
	"<tr><td align='right' height=25><b>����ȫ�ƣ�</b></td><td>&nbsp;<input class='text' size=30 name='FullName'> * </td></tr>"&_
	"<tr><td align='right' height=25><b>������ּ��</b></td><td>&nbsp;<input class='text' size=70 name='Note'> * </td></tr>"&_
	"</table><div align='center' style=""height:25px;BACKGROUND: "&BBS.SkinsPIC(2)&";""><input type='submit' class='button' value=' �� �� '>&nbsp;&nbsp;<input type='reset' class='button' value=' �� �� '></div></form>"
	BBS.ShowTable"��������",Content
Else
	IF Name="" or FullName="" or Note="" Then
		BBS.Alert"����Ҫ��д����Ϣ��û����д������","?"
	ElseIF Len(Name)>6 or Len(FullName)>50 Or Len(Note)>200 Then
		BBS.Alert"�ַ�̫�࣬��������̳�����ơ�","?"
	ElseIf int(SESSION(CacheName & "MyInfo")(6))<20 then
		BBS.Alert"����"&BBS.Info(121)&"С�� 20 ��","?"
	ElseIf int(SESSION(CacheName & "MyInfo")(7))<10000 then
		BBS.Alert"����"&BBS.Info(120)&"���� 10000 ��","?"
	ElseIf Not BBS.Execute("Select ID From[Faction] where User='"&BBS.MyName&"'").Eof Then
		BBS.Alert"���Ѿ���Ϊ�����ˣ������ٴ������ɣ�","?"
	Else
	BBS.execute("Insert into[Faction](Name,FullName,[Note],BuildDate,[User])Values('"&Name&"','"&FullName&"','"&Note&"','"&BBS.NowBbsTime&"','"&BBS.MyName&"')")
	BBS.execute("Update [User] Set Coin=Coin-10000,Faction='"&Name&"' where ID="&BBS.MyID&"")
	Session(CacheName & "MyInfo") = Empty
	BBS.Alert"���ɹ��Ĵ����˰���["&Name&"]���������Ǹð��ɵ������ˣ���ϲ����","?"
	End if
End if
End Sub

Sub Edit
Dim Name,FullName,Note,Content
Name=BBS.Fun.GetStr("Name")
FullName=BBS.Fun.GetStr("FullName")
Note=BBS.Fun.GetStr("Note")
Set Rs=BBS.Execute("Select Name,FullName,Note,User from [Faction] where ID="&ID)
If Rs.eof Then
	BBS.Alert"�����ڴ˰��ɣ�","?"
ElseIf BBS.MyName<>Rs(3) Then
	BBS.Alert"������["&Rs(0)&"]�İ����޷��޸���Ϣ��","?"
Else
	IF Name="" And FullName="" And Note="" Then
		Set Rs=BBS.Execute("Select Name,FullName,Note,User from [Faction] where ID="&ID)
		If Rs.eof Then
			BBS.Alert"�����ڴ˰��ɣ�","?"
		ElseIf BBS.MyName<>Rs(3) Then
			BBS.Alert"������["&Rs(0)&"]�İ����޷��޸���Ϣ��","?"
		Else
			Content="<form  method='post' style='margin:0'>"&_
			"<table width='95%' border=1' style='border-collapse: collapse;WORD-BREAK: break-all;'><tr><td colspan=2 height=30 align='center'><font color=red>ע�⣺ÿ���޸İ�����Ϣ���۳��� 1000 ��"&BBS.Info(120)&"�� </font></td></tr>"&_
			"<tr><td align='right' height=25><b>�������ƣ�</b></td><td>&nbsp;<input class='text' Maxlength=10 Name='Name' size='10' value='"&Rs(0)&"'>*��Ҫ����5������</td></tr>"&_
			"<tr><td align='right' height=25><b>����ȫ�ƣ�</b></td><td>&nbsp;<input class='text' size=30 name='FullName' value='"&Rs(1)&"'> * </td></tr>"&_
			"<tr><td align='right' height=25><b>������ּ��</b></td><td>&nbsp;<input class='text' size=70 name='Note' value='"&Rs(2)&"'> * </td></tr>"&_
			"</table><div align='center' style=""height:25px;BACKGROUND: "&BBS.SkinsPIC(2)&";""><input type='submit' class='button' value=' �� �� '>&nbsp;&nbsp;<input type='reset' class='button' value=' �� �� '></div></form>"
			BBS.ShowTable"��������",Content
		End If
	Else
		IF Name="" or FullName="" or Note="" Then
			BBS.Alert"����Ҫ��д����Ϣ��û����д������","?"
		ElseIf int(SESSION(CacheName & "MyInfo")(7))<1000 then
			BBS.Alert"�Բ������"&BBS.Info(120)&"����1000Ԫ���������ٰ��ɡ�","?"
		ElseIF Len(Name)>10 or Len(FullName)>50 Or Len(Note)>200 Then
			BBS.Alert"�ַ�̫�࣬��������̳�����ơ�","?"
		Else
		BBS.execute("Update [User] Set Faction='"&Name&"' where Faction='"&Rs(0)&"'")
		BBS.execute("Update [User] Set Coin=Coin-1000 where Name='"&BBS.MyName&"'")
		BBS.execute("Update [Faction]Set Name='"&Name&"',FullName='"&FullName&"',[Note]='"&Note&"' where ID="&ID)
		Session(CacheName & "MyInfo") = Empty
		BBS.Alert"�ɹ����޸��˰��ɣ�","?"
		End if
	End if
End If
Rs.Close
End Sub
%>