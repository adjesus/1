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
		Show="<div style='margin-top:9px;padding:5px'>对不起，只有本站会员才能申请友情连接！【<a style='cursor:pointer' onClick=parent.AutoLink()>关闭</a>】【<a style='cursor:pointer' onClick=location.href='register.asp'>注册</a>】【<a style='cursor:pointer' onClick=location.href='login.asp'>登陆</a>】</div>"
	Else
		Show="<div style='float:left;width=50%'><b>自助申请网站联盟说明</b><li>在贵站上先把本站的链接加入！！！</li><li>您的网站必须内容完整或制作不粗糙</li><li>不能出现任何色情、政治等国内法律不允许的内容</li><li>不能有过多弹窗、修改网友IE及修改注册表 </li><li>拒绝交换纯广告之类的网站</li><li>同意点击下一步后填写贵站相关信息</div><div><b>首页链接要求：</b><li>贵站流量在日IP在800以上</li><li>本站会员发帖量在500帖或积分20分</li><li>本站版主或会员等级在15级以上</li><li>对本站有特殊贡献的会员</li><li>以上任一项均可在首页上直接显示，</li><li>不具有上面任一项将会在收藏链接中显示</lu></div><br><div align='center'><form action='Link.asp?action=Apply' method=post style='margin:0' target='hiddenframe'><input class='BBS' type='submit' name='Submit' value=' 同 意 '>&nbsp;&nbsp;<input class='BBS' type='button' onClick=parent.AutoLink() value=' 不同意 '></form></div>"
	End If
	Show=ShowTable(Show)
End Sub

Sub Apply()
	Show="<form action='link.asp?action=savelink' method=post style='margin:0' target='hiddenframe'><li><b>请填写贵站的信息</b></li><li>论坛站长："&BBS.MyName&"</li><li>论坛名称：<input type='text' name='bbsname' size='20'></li><li>论坛地址：<input type='text' name='url' size='38' value='http://'></li><li>论坛图片：<input type='text' name='pic' size='38'> (留空则显示文字连接)</li><li>论坛说明：</td><td><input type='text' name='Readme' size='38'> (限30字内，可以留空)</li><li>图片显示：<input type='radio' name='ispic' value='yes'checked> 是 <input type='radio' name='ispic' value='no' > 否</li><br><li><input type='submit' value=' 提 交 '>&nbsp;&nbsp;<input type='reset' value=' 重 置 '></li></form>"
	Show=ShowTable(Show)
End Sub

Sub SaveLink()
	Dim BbsName,Url,Pic,Readme,Admin,Orders,IsPic
	Dim Come,Here
	Come=Request.ServerVariables("HTTP_REFERER")
	Here=Request.ServerVariables("SERVER_NAME")
	If Mid(Come,8,len(Here))<>Here then Show=ShowTable("提交失败！请不要外部提交，谢谢合作")
	BbsName=BBS.Fun.HtmlCode(BBS.Fun.GetStr("bbsname"))
	Url=BBS.Fun.HtmlCode(BBS.Fun.GetStr("url"))
	Pic=BBS.Fun.HtmlCode(BBS.Fun.GetStr("pic"))
	Readme=BBS.Fun.HtmlCode(BBS.Fun.GetStr("Readme"))
	IsPic=BBS.Fun.HtmlCode(BBS.Fun.GetStr("ispic"))
	If BbsName="" or url="" then
		Show=ShowTable("提交失败！请填写完整再提交！ 【<a style='cursor:pointer' onClick=history.go(-1)>返回重填</a>】")
		Exit Sub
	ElseIf Not BBS.Fun.CheckName(BbsName) Or (Admin<>"" And Not BBS.Fun.CheckName(Admin)) Then
		Show=ShowTable("提交失败！请不要使用了非法字符! 【<a style='cursor:pointer' onClick=history.go(-1)>返回重填</a>】")
		Exit Sub
	ElseIf Len(Readme)>30 or Len(BbsName)>15 or len(url)>250 Then
		Show=ShowTable("提交失败！字符超过了限制！ 【<a style='cursor:pointer' onClick=history.go(-1)>返回重填</a>】")
		Exit Sub
	End if
	If BBS.execute("Select admin From [Link] where Bbsname='"&BbsName&"' or url='"&Url&"' or Admin='"&BBS.MyName&"'").eof Then
			Show=ShowTable("您已经申请过了，请不要重复！ 【<a style='cursor:pointer' onClick=parent.AutoLink()>关闭</a>】")
	End If
	Orders=BBS.execute("select Count(ID) From[Link]")(0)
	Orders=Int(Orders+1)
	BBS.execute("insert into[Link](Bbsname,Url,Pic,Readme,admin,Orders,IsPic,pass)values('"&BbsName&"','"&Url&"','"&Pic&"','"&Readme&"','"&BBS.MyName&"',"&Orders&","&IsPic&",False)")
	Show=ShowTable("成功！请等待本站管理员的审核！ 【<a style='cursor:pointer' onClick=parent.AutoLink()>确定完成</a>】")
End Sub
%>

