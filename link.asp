<!--#include file="inc.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<%
Dim action,Show
action=request.querystring("action")
Select Case action
Case "Apply"
	Apply
Case "SaveLink"
	SaveLink
Case ELse
	Main
End Select
Set BBS =Nothing
%>
<script language='JavaScript'>
parent.document.getElementById("ShowAddLink").innerHTML="<%=Show%>"
</script>
<%
Function ShowTable(Str)
	ShowTable="<div  style='padding:10px;'>"&Str&"</div>"
End Function

Sub Main()
	If Not BBS.FoundUser Then
		Show="<div style='margin-top:9px;padding:5px'>�Բ���ֻ�б�վ��Ա���������������ӣ���<a style='cursor:pointer' onClick=parent.AutoLink()>�ر�</a>����<a style='cursor:pointer' onClick=location.href='register.asp'>ע��</a>����<a style='cursor:pointer' onClick=location.href='login.asp'>��½</a>��</div>"
	Else
		Show="<div style='float:left;width=50%'><b>����������վ����˵��</b><li>�ڹ�վ���Ȱѱ�վ�����Ӽ��룡����</li><li>������վ���������������������ֲ�</li><li>���ܳ����κ�ɫ�顢���εȹ��ڷ��ɲ����������</li><li>�����й��൯�����޸�����IE���޸�ע��� </li><li>�ܾ����������֮�����վ</li><li>ͬ������һ������д��վ�����Ϣ</div><div><b>��ҳ����Ҫ��</b><li>��վ��������IP��800����</li><li>��վ��Ա��������500�������20��</li><li>��վ�������Ա�ȼ���15������</li><li>�Ա�վ�����⹱�׵Ļ�Ա</li><li>������һ���������ҳ��ֱ����ʾ��</li><li>������������һ������ղ���������ʾ</lu></div><br><div align='center'><form action='Link.asp?action=Apply' method=post style='margin:0' target='hiddenframe'><input class='BBS' type='submit' name='Submit' value=' ͬ �� '>&nbsp;&nbsp;<input class='BBS' type='button' onClick=parent.AutoLink() value=' ��ͬ�� '></form></div>"
	End If
	Show=ShowTable(Show)
End Sub

Sub Apply()
	Show="<form action='link.asp?action=savelink' method=post style='margin:0' target='hiddenframe'><li><b>����д��վ����Ϣ</b></li><li>��̳վ����"&BBS.MyName&"</li><li>��̳���ƣ�<input type='text' name='bbsname' size='20'></li><li>��̳��ַ��<input type='text' name='url' size='38' value='http://'></li><li>��̳ͼƬ��<input type='text' name='pic' size='38'> (��������ʾ��������)</li><li>��̳˵����</td><td><input type='text' name='Readme' size='38'> (��30���ڣ���������)</li><li>ͼƬ��ʾ��<input type='radio' name='ispic' value='yes'checked> �� <input type='radio' name='ispic' value='no' > ��</li><br><li><input type='submit' value=' �� �� '>&nbsp;&nbsp;<input type='reset' value=' �� �� '></li></form>"
	Show=ShowTable(Show)
End Sub

Sub SaveLink()
	Dim BbsName,Url,Pic,Readme,Admin,Orders,IsPic
	Dim Come,Here
	Come=Request.ServerVariables("HTTP_REFERER")
	Here=Request.ServerVariables("SERVER_NAME")
	If Mid(Come,8,len(Here))<>Here then Show=ShowTable("�ύʧ�ܣ��벻Ҫ�ⲿ�ύ��лл����")
	BbsName=BBS.Fun.HtmlCode(BBS.Fun.GetStr("bbsname"))
	Url=BBS.Fun.HtmlCode(BBS.Fun.GetStr("url"))
	Pic=BBS.Fun.HtmlCode(BBS.Fun.GetStr("pic"))
	Readme=BBS.Fun.HtmlCode(BBS.Fun.GetStr("Readme"))
	IsPic=BBS.Fun.HtmlCode(BBS.Fun.GetStr("ispic"))
	If BbsName="" or url="" then
		Show=ShowTable("�ύʧ�ܣ�����д�������ύ�� ��<a style='cursor:pointer' onClick=history.go(-1)>��������</a>��")
		Exit Sub
	ElseIf Not BBS.Fun.CheckName(BbsName) Or (Admin<>"" And Not BBS.Fun.CheckName(Admin)) Then
		Show=ShowTable("�ύʧ�ܣ��벻Ҫʹ���˷Ƿ��ַ�! ��<a style='cursor:pointer' onClick=history.go(-1)>��������</a>��")
		Exit Sub
	ElseIf Len(Readme)>30 or Len(BbsName)>15 or len(url)>250 Then
		Show=ShowTable("�ύʧ�ܣ��ַ����������ƣ� ��<a style='cursor:pointer' onClick=history.go(-1)>��������</a>��")
		Exit Sub
	End if
	If BBS.execute("Select admin From [Link] where Bbsname='"&BbsName&"' or url='"&Url&"' or Admin='"&BBS.MyName&"'").eof Then
			Show=ShowTable("���Ѿ�������ˣ��벻Ҫ�ظ��� ��<a style='cursor:pointer' onClick=parent.AutoLink()>�ر�</a>��")
	End If
	Orders=BBS.execute("select Count(ID) From[Link]")(0)
	Orders=Int(Orders+1)
	BBS.execute("insert into[Link](Bbsname,Url,Pic,Readme,admin,Orders,IsPic,pass)values('"&BbsName&"','"&Url&"','"&Pic&"','"&Readme&"','"&BBS.MyName&"',"&Orders&","&IsPic&",False)")
	Show=ShowTable("�ɹ�����ȴ���վ����Ա����ˣ� ��<a style='cursor:pointer' onClick=parent.AutoLink()>ȷ�����</a>��")
End Sub
%>

