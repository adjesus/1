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
<title>简约论坛 - 后台管理中心</title>
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
<div style="float:right;" class=toprightdiv><a href='http://www.74177.com/bbs' target='_blank' class=topright><font color="#FF0000">官方技术支持</font></a> 欢迎您：<%=BBS.GetMemor("Admin","AdminName")%> <a href='Index.asp' target='_blank' class=topright>论坛首页</a> <a href="admin_login.asp?action=exit" target="_parent" class=topright>退出</a>
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
Menu(0,0)="admin_action.asp?action=bbsinfo,常规设置"
Menu(0,1)="?action=right,管理中心首页"
Menu(0,2)="admin_action.asp?action=bbsinfo,论坛信息设置"
Menu(0,3)="admin_action.asp?action=configdata,论坛统计设置"
Menu(0,4)="admin_actionlist.asp?action=placard,公告发布管理"
Menu(0,5)="admin_sethtmledit.asp?action=agreement,修改注册协议"
Menu(0,6)="admin_action.asp?action=gapAd,帖间广告管理"
Menu(0,7)="admin_actionlist.asp?action=link,友情链接管理"

Menu(1,0)="admin_board.asp,论坛版块"         
Menu(1,1)="admin_board.asp,论坛版面管理"
Menu(1,2)="admin_board.asp?action=addClass,添加论坛分类"       
Menu(1,3)="admin_board.asp?action=addboard,添加论坛版面"
Menu(1,4)="admin_confirm.asp?action=setjsmenu,<span style='color:#F00'>生成前台菜单</span>"

Menu(2,0)="admin_actionlist.asp?action=userlist,用户管理"
Menu(2,1)="admin_actionlist.asp?action=userlist,用户批量管理"
Menu(2,2)="admin_actionlist.asp?action=userlist&flag=2,恢复删除用户"
Menu(2,3)="admin_actionlist.asp?action=userlist&flag=1,设置 VIP用户"
Menu(2,4)="admin_actionlist.asp?action=setgrade,设置特别等级"
Menu(2,5)="admin_action.asp?action=boardadmin,设置论坛版主"
Menu(2,6)="admin_action.asp?action=grade,用户等级管理"
Menu(2,7)="admin_action.asp?action=topadmin,设置管理人员"

Menu(3,0)="admin_action.asp?action=delessay,帖子留言"         
Menu(3,1)="admin_action.asp?action=delessay,批量删除帖子"         
Menu(3,2)="admin_action.asp?action=moveessay,批量移动帖子"        
Menu(3,3)="admin_action.asp?action=delsms,批量删除留言"
Menu(3,4)="admin_sethtmledit.asp?action=allsms,群发信件留言"         
Menu(3,5)="admin_upLoad.asp,上传文件管理"
Menu(3,6)="admin_recycle.asp,论坛回收站"

Menu(4,0)="admin_new.asp,论坛插件"
Menu(4,1)="admin_new.asp,论坛调用"
Menu(4,2)="admin_action.asp?action=Bank,论坛银行管理"
Menu(4,3)="admin_action.asp?action=Faction,论坛帮派管理"

Menu(5,0)="admin_template.asp,风格模板"
Menu(5,1)="admin_template.asp,风格模板管理"
Menu(5,2)="admin_action.asp?action=Menu,论坛菜单管理"
Menu(5,3)="admin_confirm.asp?action=setjsmenu,<span style='color:#F00'>生成前台菜单</span>"

Menu(6,0)="admin_confirm.asp?action=compressdata,数据整理"
Menu(6,1)="admin_confirm.asp?action=compressdata,压缩数据库"        
Menu(6,2)="admin_confirm.asp?action=backupdata,备份数据库"        
Menu(6,3)="admin_confirm.asp?action=restoredata,恢复数据库"    
Menu(6,4)="admin_action.asp?action=sqlTable,数据表管理"
Menu(6,5)="admin_action.asp?action=updateBbs,论坛整理修复"     
Menu(6,6)="admin_user.asp?action=executesql,执行SQL语句"
Menu(6,7)="admin_action.asp?action=spacesize,空间占用情况"

Menu(7,0)="admin_actionlist.asp?action=log,系统相关"
Menu(7,1)="admin_actionlist.asp?action=log,论坛日志系统"
Menu(7,2)="admin_action.asp?action=lockip,IP封锁管理"
Menu(7,3)="admin_action.asp?action=clean,更新论坛缓存"
Menu(7,4)="admin_server.asp,服务器检测"

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
Response.Write"<div class='mian'><div class='top'>系统信息</div>"&_
"<div class='divtr1 adding'><div style='float:right;width:50%'>主题帖数："&.InfoUpdate(1)&"</div>总发帖数："&.InfoUpdate(0)&"</div>"&_
"<div class='divtr2 adding'><div style='float:right;width:50%'>昨日帖数："&.InfoUpdate(3)&"</div>今日帖数："&.InfoUpdate(2)&"</div>"&_
"<div class='divtr1 adding'><div style='float:right;width:50%'>论坛时间："&.NowBBSTime&"</div>最大日发帖数："&.InfoUpdate(4)&"</div>"&_
"<div class='divtr2 adding'><div style='float:right;width:50%'>最新会员："&.InfoUpdate(6)&"</div>会员数："&.InfoUpdate(5)&"</div>"&_
"<div class='divtr1 adding'><div style='float:right;width:50%'>最大在线人数："&.InfoUpdate(7)&"("&.InfoUpdate(8)&")</div>目前在线人数："&OnlineNum&"</div>"&_
"<div class='divtr2 adding'><div style='float:right;width:50%'>论坛版本："&.Ver&"</div>论坛访问人次："&.InfoUpdate(9)&"(超2000次更新)</div>"&_
"</div>"
Response.Write"<div class='mian'><div class='top'>快捷管理</div>"&_
"<div class='divtr1' style='padding:5px;'>【<a href='admin_Confirm.asp?action=backupdata'>数据库备份</a>】 【<a href='admin_User.asp?action=adminOK&Name="&BBS.MyName&"'><span style='color:#F00'>修改我的密码</span></a>】 【<a href='admin_confirm.asp?action=setjsmenu'><span style='color:#F00'>重建前台菜单</span></a>】 【<a href='admin_Recycle.asp'>论坛回收站</a>】 【<a href='admin_actionlist.asp?action=userlist'>用户管理</a>】<hr size=1 color=#FFFFFF />【<a href='admin_Server.asp'>服务器检测</a>】 【<a href='admin_board.asp?action=BoardUpdate'>论坛版面整理</a>】 【<a href='admin_action.asp?action=SpaceSize'>空间占用情况</a>】 【<a href='admin_actionlist.asp?action=link'><span style='color:#F00'>友情链接修改</span></a>】</div>"&_
"</div>"
Footer()
End with
End Sub

%>