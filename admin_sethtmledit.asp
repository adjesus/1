<!--#include file="Admin_Check.asp"-->
<script language="JavaScript" type="text/javascript" src="Inc/Site.js"></script>
<script language="JavaScript" type="text/javascript" src="Inc/Editor.js"></script>
<script language='JavaScript' type='text/javascript'>
function submitform(Title){
if(Title=="")Title="��û����������"
document.getElementById("content").value=IframeID.document.body.innerHTML;
if(document.getElementById("caption").value.length<1){alert(Title);document.getElementById("caption").focus();return false;};
if(document.getElementById("content").value.length<1){alert("����д�������ύ��");IframeID.focus();return false;};
form1.submit();
}
</script>
<%
Head()
Select Case Lcase(request.querystring("Action"))
Case"agreement"
	CheckString "42"
	agreement()
Case"editplacard"
	CheckString "03"
	EditPlacard()
Case"sayplacard"
	CheckString "03"
	SayPlacard()
Case"allsms"
	CheckString "34"
	AllSms
End select
Footer()

Sub agreement()
	Dim Temp,objFSO,objname
	Set objFSO = Server.CreateObject("Scr"&"ipting"&".Fil"&"eSy"&"stemOb"&"ject")
	Set objName=objFSO.OpenTextFile(Server.MapPath("inc/agreement.html"))
	Temp=objName.readall
	objName.close
	Set objFSO=nothing
	Response.Write"<form action='admin_Confirm.asp?action=saveagreement' method='post' name='form1'><input name='caption' id='caption' type='hidden' value='BBS' /><textarea id='content' name='content' style='display:none'>"&Temp&"</textarea>"&_
	"<div class='mian'><div style= 'background:#C9D4DB'><div class='top'>ע��Э��</div><div class='divtr1'><script language=""JavaScript"" type=""text/javascript"">HtmlEdit()</script></div>"&_
	"<div class='bottom'><input type='button' class='button' value='�� ��' id='sayb' onclick=submitform() /></div></div></div></form>"
End sub


Sub SayPlacard()
	Dim ID,Caption,Content,Title,name,AddTime,Hits,B_ID
	Title="��������"
	AddTime=BBS.NowBBSTime
	Name=BBS.GetMemor("Admin","AdminName")
	Hits=0
	ID=Request("ID")
	B_ID=0
	If ID<>"" Then
		Set Rs=BBS.execute("select BoardID,Caption,Content,AddTime,Name,hits from [Placard] where ID="&ID&"")
		IF Not rs.eof Then
			Title="�༭����"
			B_ID=Rs(0)
			Caption=Rs(1)
			Content=Rs(2)
			AddTime=Rs(3)
			Name=Rs(4)
			Hits=Rs(5)
		Else
			Goback"","�Ҳ�����¼�������Ѿ�ɾ���ˡ�"
		End If
	End If
	Response.Write"<form action='admin_Confirm.asp?action=SavePlacard' method='post' name='form1'><textarea id='content' name='content' style='display:none'>"&Content&"</textarea>"&_
	"<div class='mian'><div class='top'>"&Title&"</div>"
	DIVTR"������⣺","","<input name='ID' type='hidden' value='"&ID&"' /><input name='caption' type='text' class='text' id='caption' value='"&caption&"' size='60' />",25,1
	DIVTR"���ڰ�飺","","<select name='BoardID'><option value='0'>��̳��ҳ</option>"&BBS.BoardIDList(B_ID,0)&"</select>",25,2
	Response.Write"<div class='divtr2'><script language=""JavaScript"" type=""text/javascript"">HtmlEdit()</script></div>"&_
	"<div class='divtr1' style='padding:3px'>�����ߣ�<input name='Name' type='text' class='text' value='"&Name&"' size='8' />&nbsp;&nbsp;ʱ�䣺<input name='AddTime' type='text' class='text' value='"&AddTime&"' size='20' />&nbsp;&nbsp;�Ķ�������<input name='Hits' type='text' class='text' value='"&Hits&"' size='5' /></div>"&_
	"<div class='bottom'><input type='button' class='button' value='�� ��' id='sayb' onclick=submitform() /></div></div></form>"
End Sub

Sub AllSms
	Response.Write"<form action='admin_Confirm.asp?Action=AllSms' method='post' onSubmit=""ok.disabled=true;ok.value='����Ⱥ���ż�-���Եȡ�����'"" name='form1'><textarea id='content' name='content' style='display:none'></textarea>"&_
	"<div class='mian'><div style= 'background:#C9D4DB'><div class='top'>Ⱥ���ż����������û����ԣ�</div><div class='divth'>ע�⣺�˲������ܽ����Ĵ�����������Դ�������ã�</div>"
	DIVTR"�����û�Ⱥ��","","<select name='caption' style='font-size: 9pt'><option value='' selected></option><option value=1>���������û�</option><option value=7>"&BBS.GetGradeName(0,7)&"</option><option value=8>"&BBS.GetGradeName(0,8)&"</option><option value=9>"&BBS.GetGradeName(0,9)&"</option><option value=10>�����Ŷ�(����+����Ա)</option><option value=4>����Vip�û�</option><option value=0>����ע���û�(����)</option></select>",25,1
	Response.Write"<div class='divtr2'><script language=""JavaScript"" type=""text/javascript"">HtmlEdit()</script></div>"&_
	"<div class='bottom'><input type='button' name='ok' class='button' value='ȷ���ͳ�' onclick=submitform('��ѡ����յ��û�Ⱥ��') /></div></div></div></form>"
End Sub
%>
