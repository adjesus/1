<!--#include file="Admin_check.asp"-->
<script language="JavaScript" type="text/javascript">
function Show(ast){
//����
if(ast==1){
str="topic"
tmp=document.myform.bid.options[document.myform.bid.selectedIndex].value
if(tmp!="")str+="&boardid="+tmp;
tmp=document.myform.num.value;
if(tmp!="")str+="&num="+tmp;
tmp=document.myform.type.options[document.myform.type.selectedIndex].value
if(tmp!="")str+="&type="+tmp;
tmp=document.myform.order.options[document.myform.order.selectedIndex].value
if(tmp!="")str+="&order="+tmp;
tmp=document.myform.day.options[document.myform.day.selectedIndex].value
if(tmp!="")str+="&day="+tmp;
tmp=document.myform.len.value
if(tmp!="")str+="&len="+tmp;
tmp=document.myform.user.options[document.myform.user.selectedIndex].value
if(tmp!="")str+="&user="+tmp;
tmp=document.myform.time.options[document.myform.time.selectedIndex].value
if(tmp!="")str+="&time="+tmp;
tmp=document.myform.face.options[document.myform.face.selectedIndex].value
if(tmp!="")str+="&face="+tmp;
}
//��Ϣ
if(ast==2){
str="info"
var obj=document.getElementsByTagName("input");
tmp="|"
	for (var i=0;i<obj.length;i++)
	{
		if (obj[i].checked==true){tmp+=obj[i].value+"|"};
	}
	if (tmp!="|")str+="&flag="+tmp;
tmp=document.myform.face.options[document.myform.face.selectedIndex].value
if(tmp!="")str+="&face="+tmp;
}
//��Ա
if(ast==3){
str="user"
tmp=document.myform.flag.options[document.myform.flag.selectedIndex].value
if(tmp!="")str+="&flag="+tmp;
tmp=document.myform.num.value;
if(tmp!="")str+="&num="+tmp;
tmp=document.myform.face.options[document.myform.face.selectedIndex].value
if(tmp!="")str+="&face="+tmp;

}
//����
if(ast==4){
str="placard"
tmp=document.myform.bid.options[document.myform.bid.selectedIndex].value
if(tmp!="")str+="&boardid="+tmp;
tmp=document.myform.num.value;
if(tmp!="")str+="&num="+tmp;
tmp=document.myform.face.options[document.myform.face.selectedIndex].value
if(tmp!="")str+="&face="+tmp;
tmp=document.myform.time.options[document.myform.time.selectedIndex].value
if(tmp!="")str+="&time="+tmp;
tmp=document.myform.len.value
if(tmp!="")str+="&len="+tmp;
}
//���
if(ast==5){
str="board"
}
if(ast==6){
str="login"
tmp=document.myform.CK.options[document.myform.CK.selectedIndex].value
if(tmp!="")str+="&CK="+tmp;
tmp=document.myform.HI.options[document.myform.HI.selectedIndex].value
if(tmp!="")str+="&HI="+tmp;
}
//������ʽ
tmp=document.myform.tg.options[document.myform.tg.selectedIndex].value
if(tmp!="")str+="&tg="+tmp;
tmp=document.myform.h.value
if(tmp!="")str+="&h="+tmp;
tmp=document.myform.bo.options[document.myform.bo.selectedIndex].value
if(tmp!="")str+="&bo="+tmp;
tmp=document.myform.boc.value
if(tmp!="")str+="&boc="+tmp;
tmp=document.myform.bgc.value
if(tmp!="")str+="&bgc="+tmp;
document.myform.ShowScript.value='<SCR'+'IPT language="JavaScript" src="'+'<%=BBS.Info(1)%>'+'/top.asp?action='+str+'"></SC'+'RIPT>';
document.myform.ShowScript.focus();
}
function SelectColor(what){
if(!document.all){alert("��ɫ�༭�������ã���ֱ����д��ɫ���뼴�ɡ�")}
else{
	var dEL = document.all("b"+what);
	var sEL = document.all("img"+what);
	var arr = showModalDialog("pic/edit/selcolor.htm", "", "dialogWidth:18em; dialogHeight:19em; status:0;help:0;scroll:no;");
	if (arr) {
		dEL.value=arr.replace('#','');
		sEL.style.backgroundColor=arr;
	}
	}
}
</script>
<%
Head()

CheckString "09"
Response.Write"<div class='mian'><div class='top'>��̳��ҳ����</div><div class='divth'><a href='?action=topic'>�������ӵ���</a> | <a href='?action=info'>��̳��Ϣ����</a> | <a href='?action=user'>��Ա����</a> | <a href='?action=placard'>�������</a> | <a href='?action=board'>����б���</a> | <a href='?action=login'>��½��Ϣ����</a></div></div>"
Select Case Request("Action")
Case"info"
Info
Case"user"
User
Case"board"
Board
Case"placard"
Placard
Case"login"
Login
Case Else
Topic
End Select
Footer()

Sub ShowScript(ast)
Response.Write"<li>�򿪷�ʽ��<SELECT size=1 name='tg'><OPTION value=1 selected>���´��ڴ�</OPTION><OPTION value=0>�ñ����ڴ�</OPTION></SELECT></li>"&_
"<li>���߿�<SELECT size=1 name='bo'><OPTION value='' selected>0</OPTION><OPTION value=1>1</OPTION><OPTION value=1>2</OPTION></SELECT></li>"&_
"<li>ÿ�и߶ȣ�<INPUT name='h' class='text' size='2' value='18' onkeypress='if (event.keyCode < 48 || event.keyCode >  57) event.returnValue = false;' /></li>"&_
"<li>�߿���ɫ��<INPUT name='boc' class='text' size='7' /><img border='0' align=""absmiddle"" src='pic/edit/rect.gif' width='18' name='imgoc' style='cursor:pointer;' onClick=""SelectColor('oc')""> </li>"&_
"<li>��Ӱ��ɫ��<INPUT name='bgc' class='text' size='7' /><img border='0' align=""absmiddle"" src='pic/edit/rect.gif' width='18' name='imggc' style='cursor:pointer;' onClick=""SelectColor('gc')""> </li>"&_
"<li><INPUT class='button' onclick='Show("&ast&")' type='button' size='9' value='���ɵ��ô���' >��������Ĵ�����������ҳ��Ԫ���м���ʵ����̳���ӵ���</li><div style='text-align:center'><textarea name='ShowScript' rows='4'></textarea></div><div style='text-align:left; padding:5px'><b>˵����</b>�����������������ڲ����û��Լ���ͨ��ҳ�Ĵ��룬�����ܹ�����̳������Դ��̬��ʾ����ͨ��ҳ�κεط��� <br>�������ֵĴ�С��������ҳ��CSS��ʽ�����ã�</div>"
End Sub

Sub Topic
Response.Write"<div class='mian'><div class='top'>�������ӵ���</div><div class='content'><FORM name='myform' action=?type=resosave method=post>"&_
"<li style='color:#f00'>�������ã�</li>"&_
"<li>������̳��<SELECT size=1 name=bid><OPTION value=0 selected>������̳</OPTION>"&BBS.BoardIDList(0,-1)&"</SELECT></li>"&_
"<li>�������ͣ�<SELECT size=1 name='type'><OPTION selected>ȫ������</OPTION><OPTION value=1>�ö�����</OPTION><OPTION value=2>��������</OPTION><OPTION value=3>ͶƱ����</OPTION></SELECT></li>"&_
"<li>��ʾ��ʽ��<SELECT size=1 name='order'><OPTION selected>����������������</OPTION><OPTION value=1 >�����ⷢ��ʱ��</OPTION><OPTION value=2>�����ظ����⣨������</OPTION><OPTION value=3>��������������������</OPTION></SELECT></li>"&_
"<li>ʱ�䷶Χ��<SELECT size=1 name='day'><OPTION selected>��������</OPTION><OPTION value=3 >������</OPTION><OPTION value=7>һ����</OPTION><OPTION value=30>һ������</OPTION><OPTION value=90>��������</OPTION></SELECT></li>"&_
"<li>����������<INPUT name='num' class='text' onkeypress='if (event.keyCode < 48 || event.keyCode >  57) event.returnValue = false;' value='10' size=4 maxlength='2'></li>"&_
"<li>�������ƣ�<INPUT name='len' class='text' onkeypress='if (event.keyCode < 48 || event.keyCode >  57) event.returnValue = false;' value=25  size=4 maxlength='3'></li>"&_
"<li style='color:#f00'>��ʽ���ã�</li>"&_
"<li>��ͷ��ʶ��<SELECT size=1 name='face'><OPTION selected value=1>���ӱ���</OPTION><OPTION value=0 >��������</OPTION><OPTION value=*>����*</OPTION><OPTION value=��>���š�</OPTION><OPTION value=��>���š�</OPTION><OPTION value=��>���š�</OPTION><OPTION value=��>���š�</OPTION><OPTION>��Ҫ��ʶ</OPTION></SELECT></li>"&_
"<li>�������ߣ�<SELECT size=1 name='user'><OPTION value='' selected>����ʾ</OPTION><OPTION value=1>��ʾ</OPTION></SELECT></li>"&_
"<li>����ʱ�䣺<SELECT size=1 name='time'><OPTION value='' selected>����ʾ</OPTION><OPTION value=1>��ʾ</OPTION></SELECT></li>"
CALL ShowScript(1)
Response.Write"</form></div></div>"
End Sub

Sub Info
dim s,i
Response.Write"<div class='mian'><div class='top'>��̳��Ϣ����</div><div class='content'><FORM name='myform' action=?type=resosave method=post>"
s=Split("0,��̳����,��������,��������,��������,�������,ע������,���»�Ա,��̳����,���߻�Ա,�����ο�,�������,��վʱ��",",")
Response.Write"<li style='color:#f00'>������̳��Ϣ��(����������������ʾ����ѡ��)</li>"
for i=1 to uBound(s)
Response.Write"<li><input name='n"&i&"' id='n"&i&"' type='checkbox' value='"&i&"' /> "&s(i)&"</li>"
next
Response.Write"<li style='color:#f00'>��ʽ���ã�</li>"&_
"<li>��ͷ��ʶ��<SELECT size=1 name='face'><OPTION selected value='��-'>��-</OPTION></OPTION><OPTION value='*'>*</OPTION><OPTION value='��'>��</OPTION><OPTION value=��>��</OPTION><OPTION value=��>��</OPTION><OPTION value=��>��</OPTION><OPTION>��Ҫ��ʶ</OPTION></SELECT></li>"

CALL ShowScript(2)
Response.Write"</form></div></div>"
End Sub

Sub User
Response.Write"<div class='mian'><div class='top'>��̳�û�����</div><div class='content'><FORM name='myform' action=?type=resosave method=post>"&_
"<li style='color:#f00'>�����û��������ã�</li>"&_
"<li>�û����ͣ�<SELECT size=1 name='flag'><OPTION selected>������ע���������»�Ա��</OPTION><OPTION value='1'>������������򣨷����ھ���</OPTION><OPTION value='2'>�����"&BBS.Info(120)&"������̳���̣�</OPTION><OPTION value='3'>�����"&BBS.Info(121)&"����"&BBS.Info(121)&"����</OPTION><OPTION value='4'>�����"&BBS.Info(122)&"����"&BBS.Info(122)&"����</OPTION></SELECT></li>"&_
"<li>�û�������<INPUT name='num' class='text' onkeypress='if (event.keyCode < 48 || event.keyCode >  57) event.returnValue = false;' value='10' size=4 maxlength='2'></li>"&_
"<li style='color:#f00'>��ʽ���ã�</li>"&_
"<li>��ͷ��ʶ��<SELECT size=1 name='face'><OPTION selected value='��-'>��- </OPTION></OPTION><OPTION value='*'>*</OPTION><OPTION value='��'>��</OPTION><OPTION value=��>��</OPTION><OPTION value=��>��</OPTION><OPTION value=��>��</OPTION><OPTION>��Ҫ��ʶ</OPTION></SELECT></li>"
CALL ShowScript(3)
Response.Write"</form></div></div>"
End Sub

Sub Placard
Response.Write"<div class='mian'><div class='top'>��̳�������</div><div class='content'><FORM name='myform' action=?type=resosave method=post>"&_
"<li style='color:#f00'>���ù���������ã�</li>"&_
"<li>������̳��<SELECT size=1 name=bid><OPTION selected>����ȫ������</OPTION><OPTION value=0>��ҳ</OPTION>"&BBS.BoardIDList(0,-1)&"</SELECT></li>"&_
"<li>���������<INPUT name='num' class='text' onkeypress='if (event.keyCode < 48 || event.keyCode >  57) event.returnValue = false;' value='10' size=4 maxlength='2'></li>"&_
"<li>��ʾʱ�䣺<SELECT size=1 name='time'><OPTION value='' selected>����ʾ</OPTION><OPTION value=1>��ʾ</OPTION></SELECT></li>"&_
"<li style='color:#f00'>��ʽ���ã�</li>"&_
"<li>�������ƣ�<INPUT name='len' class='text' onkeypress='if (event.keyCode < 48 || event.keyCode >  57) event.returnValue = false;' value=25  size=4 maxlength='3'></li>"&_
"<li>��ͷ��ʶ��<SELECT size=1 name='face'><OPTION value=��>��</OPTION><OPTION selected value='��-'>��- </OPTION></OPTION><OPTION value='*'>*</OPTION><OPTION value='��'>��</OPTION><OPTION value=��>��</OPTION><OPTION value=��>��</OPTION><OPTION>��Ҫ��ʶ</OPTION></SELECT></li>"
CALL ShowScript(4)
Response.Write"</form></div></div>"
End Sub

Sub Login
Response.Write"<div class='mian'><div class='top'>��½������Ϣ����</div><div class='content'><FORM name='myform' action=?type=resosave method=post>"&_
"<li style='color:#f00'>�������ã�</li>"&_
"<li>Cookiesѡ�<SELECT size=1 name='CK'><OPTION value=1 selected>��ʾ</OPTION><OPTION value=''>����ʾ</OPTION></SELECT></li>"&_
"<li>��½��ʽ��<SELECT size=1 name='HI'><OPTION value=1 selected>��ʾ</OPTION><OPTION value=''>����ʾ</OPTION></SELECT></li>"&_
"<li style='color:#f00'>��ʽ���ã�</li>"
CALL ShowScript(6)
Response.Write"</form></div></div>"
End Sub

Sub board
Response.Write"<div class='mian'><div class='top'>��̳��鵼��</div><div class='content'><FORM name='myform' action=?type=resosave method=post><li style='color:#f00'>��ʽ���ã�</li>"
CALL ShowScript(5)
Response.Write"</form></div></div>"
End Sub
%>
