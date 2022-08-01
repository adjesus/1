<!--#include file="Admin_Check.asp"-->
<script language="JavaScript" type="text/javascript" src="Inc/Site.js"></script>
<script language="JavaScript" type="text/javascript" src="Inc/Editor.js"></script>
<script language='JavaScript' type='text/javascript'>
function submitform(Title){
if(Title=="")Title="还没有填完整！"
document.getElementById("content").value=IframeID.document.body.innerHTML;
if(document.getElementById("caption").value.length<1){alert(Title);document.getElementById("caption").focus();return false;};
if(document.getElementById("content").value.length<1){alert("请填写内容再提交！");IframeID.focus();return false;};
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
	"<div class='mian'><div style= 'background:#C9D4DB'><div class='top'>注册协议</div><div class='divtr1'><script language=""JavaScript"" type=""text/javascript"">HtmlEdit()</script></div>"&_
	"<div class='bottom'><input type='button' class='button' value='提 交' id='sayb' onclick=submitform() /></div></div></div></form>"
End sub


Sub SayPlacard()
	Dim ID,Caption,Content,Title,name,AddTime,Hits,B_ID
	Title="发布公告"
	AddTime=BBS.NowBBSTime
	Name=BBS.GetMemor("Admin","AdminName")
	Hits=0
	ID=Request("ID")
	B_ID=0
	If ID<>"" Then
		Set Rs=BBS.execute("select BoardID,Caption,Content,AddTime,Name,hits from [Placard] where ID="&ID&"")
		IF Not rs.eof Then
			Title="编辑公告"
			B_ID=Rs(0)
			Caption=Rs(1)
			Content=Rs(2)
			AddTime=Rs(3)
			Name=Rs(4)
			Hits=Rs(5)
		Else
			Goback"","找不到记录，可能已经删除了。"
		End If
	End If
	Response.Write"<form action='admin_Confirm.asp?action=SavePlacard' method='post' name='form1'><textarea id='content' name='content' style='display:none'>"&Content&"</textarea>"&_
	"<div class='mian'><div class='top'>"&Title&"</div>"
	DIVTR"公告标题：","","<input name='ID' type='hidden' value='"&ID&"' /><input name='caption' type='text' class='text' id='caption' value='"&caption&"' size='60' />",25,1
	DIVTR"所在版块：","","<select name='BoardID'><option value='0'>论坛首页</option>"&BBS.BoardIDList(B_ID,0)&"</select>",25,2
	Response.Write"<div class='divtr2'><script language=""JavaScript"" type=""text/javascript"">HtmlEdit()</script></div>"&_
	"<div class='divtr1' style='padding:3px'>发布者：<input name='Name' type='text' class='text' value='"&Name&"' size='8' />&nbsp;&nbsp;时间：<input name='AddTime' type='text' class='text' value='"&AddTime&"' size='20' />&nbsp;&nbsp;阅读次数：<input name='Hits' type='text' class='text' value='"&Hits&"' size='5' /></div>"&_
	"<div class='bottom'><input type='button' class='button' value='提 交' id='sayb' onclick=submitform() /></div></div></form>"
End Sub

Sub AllSms
	Response.Write"<form action='admin_Confirm.asp?Action=AllSms' method='post' onSubmit=""ok.disabled=true;ok.value='正在群发信件-请稍等。。。'"" name='form1'><textarea id='content' name='content' style='display:none'></textarea>"&_
	"<div class='mian'><div style= 'background:#C9D4DB'><div class='top'>群发信件（批量给用户留言）</div><div class='divth'>注意：此操作可能将消耗大量服务器资源。请慎用！</div>"
	DIVTR"接收用户群：","","<select name='caption' style='font-size: 9pt'><option value='' selected></option><option value=1>所有在线用户</option><option value=7>"&BBS.GetGradeName(0,7)&"</option><option value=8>"&BBS.GetGradeName(0,8)&"</option><option value=9>"&BBS.GetGradeName(0,9)&"</option><option value=10>管理团队(版主+管理员)</option><option value=4>所有Vip用户</option><option value=0>所有注册用户(慎用)</option></select>",25,1
	Response.Write"<div class='divtr2'><script language=""JavaScript"" type=""text/javascript"">HtmlEdit()</script></div>"&_
	"<div class='bottom'><input type='button' name='ok' class='button' value='确定送出' onclick=submitform('请选择接收的用户群！') /></div></div></div></form>"
End Sub
%>
