<!--#include file="admin_check.asp"-->
<%
If Request("action")="right" Then
	Call MainRight()
Else
	Call main()
End If
Set BBS =Nothing

Sub Main()
%>
<html>
<head>
<title>��Լ��̳ - ��̨��������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
.MenuT{float:left;cursor:pointer;padding:2px 5px 2px 5px;background:#3A6592;color:#FFFFFF;margin:3px 5px 3px 5px;}
.MenuT1{float:left;cursor:pointer;padding:2px 5px 2px 5px;background:#C4D8ED;color:#135294;margin:3px 5px 3px 5px;}
.MenuT2{float:left;cursor:pointer;padding:2px 5px 2px 5px;background:#4877A9;color:#FFFFFF;margin:3px 5px 3px 5px;}
.admintop{border:0px;background: #3A6592;height:20px;color:#FFFFFF}
.toprightdiv{padding:6px}
A.topright{COLOR: #FFFFFF; TEXT-DECORATION: None}
A.topright:link	{COLOR: #FFFFFF; TEXT-DECORATION: None}
A.topright:visited{COLOR: #FFFFFF; TEXT-DECORATION: None}
A.topright:hover{COLOR: #FFFFFF; TEXT-DECORATION: None}
A.topright:active{TEXT-DECORATION: none}
-->
</style>
<script language="javaScript" src="inc/Site.js" type="text/javascript"></script>
<script language=javascript>
function s(str,num){
  for (var i=0;i<=num;i++)    {
    document.getElementById("t"+i).className='MenuT';
  }
  str.className='MenuT1';
}
function m(str,num){
  for (var i=0;i<=num;i++)    {
    if(document.getElementById("t"+i).className!='MenuT1'){
	  document.getElementById("t"+i).className='MenuT';
	}
  }
  if(str.className!='MenuT1'){
    str.className='MenuT2';
  }
}
</script>
</head>
<body scroll="no" style="MARGIN: 0px">
<table width="99%" height="100%" border="0" cellspacing="0" cellpadding="0">
  <tr><td class=admintop>
<div style="float:right;" class=toprightdiv><a href='http://www.74177.com/bbs' target='_blank' class=topright><font color="#FF0000">�ٷ�����֧��</font></a> ��ӭ����<%=BBS.GetMemor("Admin","AdminName")%> <a href='Index.asp' target='_blank' class=topright>��̳��ҳ</a> <a href="admin_login.asp?action=exit" target="_parent" class=topright>�˳�</a>
</div>
<div style="float:left;width:100px"><a href='admin_index.asp'><img src=images/icon/logo.gif align=absmiddle border=0 /></a></div>
<div style="float:left;"><%Call AdminMenu()%></div>
  </td></tr>
  <tr><td><iframe id="Right" name="Right" scrolling=yes style="HEIGHT: 100%; WIDTH: 100%; Z-INDEX: 1" frameborder="0" src="?action=right"></iframe></td></tr>
</table>
</body>
</html>

<%End Sub

Sub AdminMenu()
Dim I,II,Menu(7,7),menuUrl,MenuName,Temp,TempList
Menu(0,0)="admin_action.asp?action=bbsinfo,��������"
Menu(0,1)="?action=right,����������ҳ"
Menu(0,2)="admin_action.asp?action=bbsinfo,��̳��Ϣ����"
Menu(0,3)="admin_action.asp?action=configdata,��̳ͳ������"
Menu(0,4)="admin_actionlist.asp?action=placard,���淢������"
Menu(0,5)="admin_sethtmledit.asp?action=agreement,�޸�ע��Э��"
Menu(0,6)="admin_action.asp?action=gapAd,���������"
Menu(0,7)="admin_actionlist.asp?action=link,�������ӹ���"

Menu(1,0)="admin_board.asp,��̳���"         
Menu(1,1)="admin_board.asp,��̳�������"
Menu(1,2)="admin_board.asp?action=addClass,�����̳����"       
Menu(1,3)="admin_board.asp?action=addboard,�����̳����"
Menu(1,4)="admin_confirm.asp?action=setjsmenu,<span style='color:#F00'>����ǰ̨�˵�</span>"

Menu(2,0)="admin_actionlist.asp?action=userlist,�û�����"
Menu(2,1)="admin_actionlist.asp?action=userlist,�û���������"
Menu(2,2)="admin_actionlist.asp?action=userlist&flag=2,�ָ�ɾ���û�"
Menu(2,3)="admin_actionlist.asp?action=userlist&flag=1,���� VIP�û�"
Menu(2,4)="admin_actionlist.asp?action=setgrade,�����ر�ȼ�"
Menu(2,5)="admin_action.asp?action=boardadmin,������̳����"
Menu(2,6)="admin_action.asp?action=grade,�û��ȼ�����"
Menu(2,7)="admin_action.asp?action=topadmin,���ù�����Ա"

Menu(3,0)="admin_action.asp?action=delessay,��������"         
Menu(3,1)="admin_action.asp?action=delessay,����ɾ������"         
Menu(3,2)="admin_action.asp?action=moveessay,�����ƶ�����"        
Menu(3,3)="admin_action.asp?action=delsms,����ɾ������"
Menu(3,4)="admin_sethtmledit.asp?action=allsms,Ⱥ���ż�����"         
Menu(3,5)="admin_upLoad.asp,�ϴ��ļ�����"
Menu(3,6)="admin_recycle.asp,��̳����վ"

Menu(4,0)="admin_new.asp,��̳���"
Menu(4,1)="admin_new.asp,��̳����"
Menu(4,2)="admin_action.asp?action=Bank,��̳���й���"
Menu(4,3)="admin_action.asp?action=Faction,��̳���ɹ���"

Menu(5,0)="admin_template.asp,���ģ��"
Menu(5,1)="admin_template.asp,���ģ�����"
Menu(5,2)="admin_action.asp?action=Menu,��̳�˵�����"
Menu(5,3)="admin_confirm.asp?action=setjsmenu,<span style='color:#F00'>����ǰ̨�˵�</span>"

Menu(6,0)="admin_confirm.asp?action=compressdata,��������"
Menu(6,1)="admin_confirm.asp?action=compressdata,ѹ�����ݿ�"        
Menu(6,2)="admin_confirm.asp?action=backupdata,�������ݿ�"        
Menu(6,3)="admin_confirm.asp?action=restoredata,�ָ����ݿ�"    
Menu(6,4)="admin_action.asp?action=sqlTable,���ݱ����"
Menu(6,5)="admin_action.asp?action=updateBbs,��̳�����޸�"     
Menu(6,6)="admin_user.asp?action=executesql,ִ��SQL���"
Menu(6,7)="admin_action.asp?action=spacesize,�ռ�ռ�����"

Menu(7,0)="admin_actionlist.asp?action=log,ϵͳ���"
Menu(7,1)="admin_actionlist.asp?action=log,��̳��־ϵͳ"
Menu(7,2)="admin_action.asp?action=lockip,IP��������"
Menu(7,3)="admin_action.asp?action=clean,������̳����"
Menu(7,4)="admin_server.asp,���������"

For i=0 to ubound(menu,1)
If isempty(menu(i,1)) then exit for
 Temp = "<div class="
 If i=0 Then Temp = Temp & "MenuT1" Else Temp = Temp & "MenuT"
 Temp = Temp & " id=t"&i&" onmouseover=""m(t"&i&","&ubound(menu,1)&");dropdownmenu(this, event, 'M"&i&"');"" onclick=""s(t"&i&","&ubound(menu,1)&");Right.location = '"&split(Menu(i,0),",")(0)&"'"">"&split(Menu(i,0),",")(1)&"</div>"
 Temp = Temp & "<DIV id=M"&i&" class=menu>"
   For II=1 to ubound(menu,2)
    If isempty(menu(I,II)) then Exit for
      MenuUrl=Split(menu(I,II),",")(0)
      MenuName=Split(menu(I,II),",")(1)
      Temp = Temp & "<div class=menuitems><A href="&MenuUrl&" target=Right onMouseDown=""s(t"&i&","&ubound(menu,1)&");"">"&MenuName&"</A></div>"
   Next
 Temp = Temp & "</DIV>"
 TempList = TempList & Temp
Next
Response.write TempList
End Sub

Sub MainRight()
Dim Temp,OnlineNum
with BBS
If .Cache.valid("OnlineCache") Then
	Temp=.Cache.Value("OnlineCache")
	Temp=Split(Temp,",")
	OnlineNum=uBound(Temp)+1
Else
	OnlineNum=1
End If
Response.Write"<div class='mian'><div class='top'>ϵͳ��Ϣ</div>"&_
"<div class='divtr1 adding'><div style='float:right;width:50%'>����������"&.InfoUpdate(1)&"</div>�ܷ�������"&.InfoUpdate(0)&"</div>"&_
"<div class='divtr2 adding'><div style='float:right;width:50%'>����������"&.InfoUpdate(3)&"</div>����������"&.InfoUpdate(2)&"</div>"&_
"<div class='divtr1 adding'><div style='float:right;width:50%'>��̳ʱ�䣺"&.NowBBSTime&"</div>����շ�������"&.InfoUpdate(4)&"</div>"&_
"<div class='divtr2 adding'><div style='float:right;width:50%'>���»�Ա��"&.InfoUpdate(6)&"</div>��Ա����"&.InfoUpdate(5)&"</div>"&_
"<div class='divtr1 adding'><div style='float:right;width:50%'>�������������"&.InfoUpdate(7)&"("&.InfoUpdate(8)&")</div>Ŀǰ����������"&OnlineNum&"</div>"&_
"<div class='divtr2 adding'><div style='float:right;width:50%'>��̳�汾��"&.Ver&"</div>��̳�����˴Σ�"&.InfoUpdate(9)&"(��2000�θ���)</div>"&_
"</div>"
Response.Write"<div class='mian'><div class='top'>��ݹ���</div>"&_
"<div class='divtr1' style='padding:5px;'>��<a href='admin_Confirm.asp?action=backupdata'>���ݿⱸ��</a>�� ��<a href='admin_User.asp?action=adminOK&Name="&BBS.MyName&"'><span style='color:#F00'>�޸��ҵ�����</span></a>�� ��<a href='admin_confirm.asp?action=setjsmenu'><span style='color:#F00'>�ؽ�ǰ̨�˵�</span></a>�� ��<a href='admin_Recycle.asp'>��̳����վ</a>�� ��<a href='admin_actionlist.asp?action=userlist'>�û�����</a>��<hr size=1 color=#FFFFFF />��<a href='admin_Server.asp'>���������</a>�� ��<a href='admin_board.asp?action=BoardUpdate'>��̳��������</a>�� ��<a href='admin_action.asp?action=SpaceSize'>�ռ�ռ�����</a>�� ��<a href='admin_actionlist.asp?action=link'><span style='color:#F00'>���������޸�</span></a>��</div>"&_
"</div>"
Footer()
End with
End Sub

%>