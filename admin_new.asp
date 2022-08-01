<!--#include file="Admin_check.asp"-->
<script language="JavaScript" type="text/javascript">
function Show(ast){
//主题
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
//信息
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
//会员
if(ast==3){
str="user"
tmp=document.myform.flag.options[document.myform.flag.selectedIndex].value
if(tmp!="")str+="&flag="+tmp;
tmp=document.myform.num.value;
if(tmp!="")str+="&num="+tmp;
tmp=document.myform.face.options[document.myform.face.selectedIndex].value
if(tmp!="")str+="&face="+tmp;

}
//公告
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
//版块
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
//共用样式
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
if(!document.all){alert("颜色编辑器不可用，请直接填写颜色代码即可。")}
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
Response.Write"<div class='mian'><div class='top'>论坛首页调用</div><div class='divth'><a href='?action=topic'>主题帖子调用</a> | <a href='?action=info'>论坛信息调用</a> | <a href='?action=user'>会员调用</a> | <a href='?action=placard'>公告调用</a> | <a href='?action=board'>版块列表导航</a> | <a href='?action=login'>登陆信息调用</a></div></div>"
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
Response.Write"<li>打开方式：<SELECT size=1 name='tg'><OPTION value=1 selected>用新窗口打开</OPTION><OPTION value=0>用本窗口打开</OPTION></SELECT></li>"&_
"<li>表格边框：<SELECT size=1 name='bo'><OPTION value='' selected>0</OPTION><OPTION value=1>1</OPTION><OPTION value=1>2</OPTION></SELECT></li>"&_
"<li>每行高度：<INPUT name='h' class='text' size='2' value='18' onkeypress='if (event.keyCode < 48 || event.keyCode >  57) event.returnValue = false;' /></li>"&_
"<li>边框颜色：<INPUT name='boc' class='text' size='7' /><img border='0' align=""absmiddle"" src='pic/edit/rect.gif' width='18' name='imgoc' style='cursor:pointer;' onClick=""SelectColor('oc')""> </li>"&_
"<li>背影颜色：<INPUT name='bgc' class='text' size='7' /><img border='0' align=""absmiddle"" src='pic/edit/rect.gif' width='18' name='imggc' style='cursor:pointer;' onClick=""SelectColor('gc')""> </li>"&_
"<li><INPUT class='button' onclick='Show("&ast&")' type='button' size='9' value='生成调用代码' >↓把下面的代码插入你的网页单元格中即可实现论坛帖子调用</li><div style='text-align:center'><textarea name='ShowScript' rows='4'></textarea></div><div style='text-align:left; padding:5px'><b>说明：</b>本功能用于生成用于插入用户自己普通网页的代码，代码能够把论坛主题资源动态显示在普通网页任何地方！ <br>至于文字的大小将跟随主页的CSS样式表设置！</div>"
End Sub

Sub Topic
Response.Write"<div class='mian'><div class='top'>主题帖子调用</div><div class='content'><FORM name='myform' action=?type=resosave method=post>"&_
"<li style='color:#f00'>参数设置：</li>"&_
"<li>调用论坛：<SELECT size=1 name=bid><OPTION value=0 selected>所有论坛</OPTION>"&BBS.BoardIDList(0,-1)&"</SELECT></li>"&_
"<li>主题类型：<SELECT size=1 name='type'><OPTION selected>全部主题</OPTION><OPTION value=1>置顶主题</OPTION><OPTION value=2>精华主题</OPTION><OPTION value=3>投票主题</OPTION></SELECT></li>"&_
"<li>显示方式：<SELECT size=1 name='order'><OPTION selected>按最后更新主题排列</OPTION><OPTION value=1 >按主题发布时间</OPTION><OPTION value=2>按最多回复主题（热帖）</OPTION><OPTION value=3>按主题点击数（人气帖）</OPTION></SELECT></li>"&_
"<li>时间范围：<SELECT size=1 name='day'><OPTION selected>所有日期</OPTION><OPTION value=3 >三天内</OPTION><OPTION value=7>一周内</OPTION><OPTION value=30>一个月内</OPTION><OPTION value=90>三个月内</OPTION></SELECT></li>"&_
"<li>主题数量：<INPUT name='num' class='text' onkeypress='if (event.keyCode < 48 || event.keyCode >  57) event.returnValue = false;' value='10' size=4 maxlength='2'></li>"&_
"<li>字数限制：<INPUT name='len' class='text' onkeypress='if (event.keyCode < 48 || event.keyCode >  57) event.returnValue = false;' value=25  size=4 maxlength='3'></li>"&_
"<li style='color:#f00'>样式设置：</li>"&_
"<li>题头标识：<SELECT size=1 name='face'><OPTION selected value=1>帖子表情</OPTION><OPTION value=0 >数字序列</OPTION><OPTION value=*>符号*</OPTION><OPTION value=★>符号★</OPTION><OPTION value=☆>符号☆</OPTION><OPTION value=◆>符号◆</OPTION><OPTION value=◇>符号◇</OPTION><OPTION>不要标识</OPTION></SELECT></li>"&_
"<li>帖子作者：<SELECT size=1 name='user'><OPTION value='' selected>不显示</OPTION><OPTION value=1>显示</OPTION></SELECT></li>"&_
"<li>发帖时间：<SELECT size=1 name='time'><OPTION value='' selected>不显示</OPTION><OPTION value=1>显示</OPTION></SELECT></li>"
CALL ShowScript(1)
Response.Write"</form></div></div>"
End Sub

Sub Info
dim s,i
Response.Write"<div class='mian'><div class='top'>论坛信息调用</div><div class='content'><FORM name='myform' action=?type=resosave method=post>"
s=Split("0,论坛帖数,主题帖数,今日帖数,昨日帖数,最高日帖,注册人数,最新会员,论坛在线,在线会员,在线游客,最高在线,建站时间",",")
Response.Write"<li style='color:#f00'>调用论坛信息：(下面各项如果不想显示，请选择)</li>"
for i=1 to uBound(s)
Response.Write"<li><input name='n"&i&"' id='n"&i&"' type='checkbox' value='"&i&"' /> "&s(i)&"</li>"
next
Response.Write"<li style='color:#f00'>样式设置：</li>"&_
"<li>题头标识：<SELECT size=1 name='face'><OPTION selected value='□-'>□-</OPTION></OPTION><OPTION value='*'>*</OPTION><OPTION value='★'>★</OPTION><OPTION value=☆>☆</OPTION><OPTION value=◆>◆</OPTION><OPTION value=◇>◇</OPTION><OPTION>不要标识</OPTION></SELECT></li>"

CALL ShowScript(2)
Response.Write"</form></div></div>"
End Sub

Sub User
Response.Write"<div class='mian'><div class='top'>论坛用户调用</div><div class='content'><FORM name='myform' action=?type=resosave method=post>"&_
"<li style='color:#f00'>调用用户参数设置：</li>"&_
"<li>用户类型：<SELECT size=1 name='flag'><OPTION selected>按最新注册排序（最新会员）</OPTION><OPTION value='1'>按最多帖数排序（发帖冠军）</OPTION><OPTION value='2'>按最多"&BBS.Info(120)&"排序（论坛富翁）</OPTION><OPTION value='3'>按最多"&BBS.Info(121)&"排序（"&BBS.Info(121)&"王）</OPTION><OPTION value='4'>按最多"&BBS.Info(122)&"排序（"&BBS.Info(122)&"王）</OPTION></SELECT></li>"&_
"<li>用户数量：<INPUT name='num' class='text' onkeypress='if (event.keyCode < 48 || event.keyCode >  57) event.returnValue = false;' value='10' size=4 maxlength='2'></li>"&_
"<li style='color:#f00'>样式设置：</li>"&_
"<li>题头标识：<SELECT size=1 name='face'><OPTION selected value='□-'>□- </OPTION></OPTION><OPTION value='*'>*</OPTION><OPTION value='★'>★</OPTION><OPTION value=☆>☆</OPTION><OPTION value=◆>◆</OPTION><OPTION value=◇>◇</OPTION><OPTION>不要标识</OPTION></SELECT></li>"
CALL ShowScript(3)
Response.Write"</form></div></div>"
End Sub

Sub Placard
Response.Write"<div class='mian'><div class='top'>论坛公告调用</div><div class='content'><FORM name='myform' action=?type=resosave method=post>"&_
"<li style='color:#f00'>调用公告参数设置：</li>"&_
"<li>调用论坛：<SELECT size=1 name=bid><OPTION selected>调用全部公告</OPTION><OPTION value=0>首页</OPTION>"&BBS.BoardIDList(0,-1)&"</SELECT></li>"&_
"<li>最大数量：<INPUT name='num' class='text' onkeypress='if (event.keyCode < 48 || event.keyCode >  57) event.returnValue = false;' value='10' size=4 maxlength='2'></li>"&_
"<li>显示时间：<SELECT size=1 name='time'><OPTION value='' selected>不显示</OPTION><OPTION value=1>显示</OPTION></SELECT></li>"&_
"<li style='color:#f00'>样式设置：</li>"&_
"<li>字数限制：<INPUT name='len' class='text' onkeypress='if (event.keyCode < 48 || event.keyCode >  57) event.returnValue = false;' value=25  size=4 maxlength='3'></li>"&_
"<li>题头标识：<SELECT size=1 name='face'><OPTION value=◆>◆</OPTION><OPTION selected value='□-'>□- </OPTION></OPTION><OPTION value='*'>*</OPTION><OPTION value='★'>★</OPTION><OPTION value=☆>☆</OPTION><OPTION value=◇>◇</OPTION><OPTION>不要标识</OPTION></SELECT></li>"
CALL ShowScript(4)
Response.Write"</form></div></div>"
End Sub

Sub Login
Response.Write"<div class='mian'><div class='top'>登陆窗口信息调用</div><div class='content'><FORM name='myform' action=?type=resosave method=post>"&_
"<li style='color:#f00'>参数设置：</li>"&_
"<li>Cookies选项：<SELECT size=1 name='CK'><OPTION value=1 selected>显示</OPTION><OPTION value=''>不显示</OPTION></SELECT></li>"&_
"<li>登陆方式：<SELECT size=1 name='HI'><OPTION value=1 selected>显示</OPTION><OPTION value=''>不显示</OPTION></SELECT></li>"&_
"<li style='color:#f00'>样式设置：</li>"
CALL ShowScript(6)
Response.Write"</form></div></div>"
End Sub

Sub board
Response.Write"<div class='mian'><div class='top'>论坛版块导航</div><div class='content'><FORM name='myform' action=?type=resosave method=post><li style='color:#f00'>样式设置：</li>"
CALL ShowScript(5)
Response.Write"</form></div></div>"
End Sub
%>
