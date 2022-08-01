<!--#include file="Inc.asp"-->
<!--#include file="Inc/md5.asp"-->
<%
Dim action
action=Lcase(Request.Querystring("action"))
If Len(action)>15 Then BBS.GotoErr(1)
BBS.CheckMake
if action<>"getpassword" and action<>"forgetpassword" Then
	BBS.Position=BBS.Position&" -> <a href=""userinfo.asp"">用户控制面版<a>"
	If Not BBS.FoundUser Then BBS.GotoErr(4)
End If

Select Case action
Case"myinfo"
	Myinfo()
Case"savemyinfo"
	savemyinfo
Case"mypassword"
	MyPassword
Case"savemypassword"
	savemypassword
Case"forgetpassword"
	forgetpassword
Case"getpassword"
	GetPassword
Case else
	BBS.GoToErr(1)
End Select
BBS.Footer()
Set BBS =Nothing
Sub MyManager()
	Response.Write BBS.ReadSkins("用户控制面版")
End Sub

Sub Myinfo()
	Dim Rs,S,temp,Temp1
	BBS.Head "","","修改个人资料"
	MyManager()
	SET RS=BBS.Execute("Select Name,Sex,Birthday,Mail,Home,IsQQpic,QQ,Pic,Pich,Picw,IsSign,Sign,Honor,GradeID From[user]where ID="&BBS.MyID&" And Isdel=False")
	IF Rs.eof Then BBS.MakeCookiesEmpty():BBS.GotoErr(100)
	S="<form style='margin:0' method='POST' action='?action=savemyinfo'  name='form'>"
	S=S&"<div style='text-align:left;border-bottom:1px solid "&BBS.SkinsPIC(0)&";padding:3px;font-weight:bold;BACKGROUND: "&BBS.SkinsPIC(2)&";'><img src='Images/icon/inn.gif' align='absmiddle' atl='必填资料' /> 必填资料</div>"
	S=S&BBS.Row("<b>用户名称</b>：<br>此为论坛ID帐号，自己不能修改","<br /><b>"&Rs("Name")&"</b>","65%","45px")
	If Rs("Sex")=1 Then Temp=" checked" Else Temp=""
	If Rs("Sex")=0 Then Temp1=" checked" Else Temp1=""
	S=S&BBS.Row("<b>您的性别：</b>","<input name='sex' type='radio' value='1' "&Temp&" class=checkbox /><img src='Images/icon/male.gif' align='absmiddle' atl='帅哥' /> 帅哥&nbsp;&nbsp;<input type='radio' name='sex' value='0' "&Temp1&" class=checkbox /><img src='Images/icon/female.gif' align='absmiddle' atl='靓女' /> 靓女","65%","")
	S=S&BBS.Row("<b>Email地址</b>：<br>请输入有效的邮件地址","<input type='text' class='text' name='mail' style='margin-top:8px' size='30' maxlength='30' value='"&Rs("Mail")&"' />","65%","44px")
	S=S&"<div style='text-align:left;border-bottom:1px solid "&BBS.SkinsPIC(0)&";padding:3px;font-weight:bold;BACKGROUND: "&BBS.SkinsPIC(2)&";'><img src='Images/icon/inn.gif' align='absmiddle' atl='选填资料' /> 选填资料</div>"
	S=S&"<div id='Rowhow' display='none'>"
	S=S&BBS.Row("<b>生日：</b>","<input  type='text' class='text' name='birthday' id='birthday' size='10' maxlength='10' readonly='true' onfocus=""show_cele_date('birthday','','',this)"" value='"&Rs("Birthday")&"' />","65%","20px")
	S=S&BBS.Row("<b>主页：</b><br />填写你的个人主页，让大家见识见识！","<input type='text' class='text' name='home' size='30' maxlength='200'  style='margin-top:8px' value='"&Rs("Home")&"' />","65%","44px")
	If Rs("IsQQPic") Then
	Temp=" checked"
	Else
	Temp=""
	End If
	S=S&BBS.Row("<b>QQ号码：</b><br />填写您的QQ地址，方便与他人的联系","<input type='text' class='text'  style='margin-top:8px' value='"&Rs("QQ")&"'name='QQ'  maxlength='15'> <input type='checkbox' onclick='QQpic()' name='isqqpic' id='isqqpic' value='1' "&Temp&" class=checkbox />启用QQ形象作为头像","65%","44px")
	S=S&"<div id='showpic'>"
	S=S&BBS.Row("<b>选择论坛头像：</b><br />使用论坛自带的图像",HeadPicOpt() &"<img src='"&Rs("Pic")&"' id='pic' name='pic' /> <input onclick=""openwin('preview.asp?action=HeadPic',680,400,'yes')"" type='button' class='button' value='全部头像' />","65%","50")
	If SESSION(CacheName& "MyGradeinfo")(14)="1" then
	Temp="<input type=button value='上传头像图片' class='button' onclick=""javascript:up.style.display='block';upf.location.href='UploadFile.asp?Flag=1';this.style.display='none'""><div id='up' style='display:none'><iframe id='upf' scrolling='no' frameborder='0' height='22' width='100%'></iframe></div><br>"
	Else
	Temp=""
	End if
	S=S&BBS.Row("<b>自定义头像：</b><br />如果图像位置中有连接图片将以自定义的为主",Temp&"<input id='picurl' name='picurl' size='40' maxlength='100'  value='"&Rs("Pic")&"' /> 完整Url地址<br />图像宽度：<input type='text' class='text' name='picw' id='picw'  size='6' value='"&Rs("PicW")&"' /> 高度：<input type=text name='pich' id='pich' size='6'  value='"&Rs("PicH")&"' />(最大限度:120)","65%","")
	S=S&"</div>"
	If SESSION(CacheName& "MyGradeinfo")(8)="0" then
		 Temp="您还没达到可以自定头衔称号的等级权限<input type='hidden' value='"&Rs("Honor")&"' name='Honor'>"
	Else
		 Temp="<input style='margin-top:8px' type='text' value='"&Rs("Honor")&"' name='Honor'>"
	End If
	S=S&BBS.Row("<b>自定义头衔称号：</b><br />最多8个汉字",Temp,"65%","40px")
	S=S&BBS.Row("<b>个性签名：</b><br />文字将出现在您发表的文章的结尾处<br />体现您的个性(最多255个字符)","<TEXTAREA name='sign' rows='4' style='width: 98%;'>"&Rs("Sign")&"</textarea>","65%","")
	S=S&"</div><div style=""padding:5px;BACKGROUND: "&BBS.SkinsPIC(2)&";""  align='center'><input type='submit' value='确定修改！' class='button' />&nbsp;&nbsp; <input type='reset' class='button' value='取消重写！'  /></div></form>"
	BBS.ShowTable"修改我的资料",S
%>
<script language="JavaScript" type="text/javascript" src="Inc/pswdplc.js"></script>
<script language="JavaScript" type="text/javascript" src="Inc/InputDate.js"></script>
<script type="text/javascript">
QQpic()
function QQpic(){
if (document.getElementById("isqqpic").checked == true){
	document.getElementById("showpic").style.display="none";
	}else{
	document.getElementById("showpic").style.display="block";
	}
	}
function ShowPic(){
document.getElementById("pic").src="pic/headpic/"+document.getElementById("headpicoption").options[document.getElementById("headpicoption").selectedIndex].value+".gif";
document.getElementById("picurl").value="pic/headpic/"+document.getElementById("headpicoption").options[document.getElementById("headpicoption").selectedIndex].value+".gif";
document.getElementById("pich").value='<%=BBS.info(55)%>';
document.getElementById("picw").value='<%=BBS.info(54)%>';
}
function Check(){
openwin("preview.asp?action=CheckName&name="+document.getElementById("name").value,300,30,"no");
}
</script>
<%
End Sub

	
Function HeadPicOpt()
	Dim Temp,i
	for i=2 to Int(BBS.info(53))
		Temp=Temp&"<option value='"&i&"'>"&i&"</option>"
	Next
	HeadPicOpt="<select name='headpicoption' id='headpicoption' onChange='ShowPic()'><option value='1' selected>1</option>"&Temp&"</select>"
End Function	


Sub savemyinfo()
With BBS
	.Head"","","修改个人资料"
	MyManager()
	Dim Temp,Content,Rs,Sql,Name,Mail,PicUrl,HeadPic,PicW,PicH,Home,Sign,QQ,IsQQpic,Sex,Birthday,Honor
	Name=.Fun.GetStr("name")
	Mail=.Fun.GetStr("mail")
	Honor=.Fun.GetStr("Honor")
	If Mail="" Then .GoToErr(42)
	Mail=server.HTMLEnCode(Mail)
	If Not .Fun.IsValidEmail(Mail) Then .GoToErr(42)
	'只允一个邮箱
	If .info(42)="1" Then
		If Not .Execute("SELECT ID FROM [user] where Mail='"&Mail&"' And ID<>"&.MyID).Eof Then .GoToErr(49)
	End If
	PicUrl=lcase(.Fun.HtmlCode(.Fun.GetStr("PicUrl")))
	headpic=.Fun.HtmlCode(.Fun.GetStr("headpicoption"))
	If Not .Fun.isInteger(headpic) Or Not .Fun.IsUrl(PicUrl) Then .GoToErr(81)
	Home=.Fun.HtmlCode(.Fun.GetStr("Home"))
	Sex=.Fun.GetStr("Sex")
	Birthday=.Fun.GetStr("Birthday")
	QQ=.Fun.GetStr("QQ")
	IsQQpic=.Fun.GetStr("IsQQpic")
	If IsQQPic="" Then IsQQPic=0
	If Instr(Home,"://")=0 Then Home=.info(1)
	If IsQQpic<>"1" Then IsQQpic="0"
	Sign=Replace(Left(.Fun.Replacehtml(.Fun.GetStr("Sign")),255),"{帖子内容}","")
	PicH=.Fun.GetStr("pich")
	PicW=.Fun.Getstr("picw")
	If .info(57)="1" And (Instr(PicUrl,"://")>0  Or Instr(Lcase(Picurl),"www")>0 Or Instr(Lcase(PicUrl),"..")>0) Then  .GotoErr(45)'禁止外部图片
	If PicUrl="" then
		PicUrl="Pic/headpic/"& HeadPic &".gif"
		PicW= .info(54)
		PicH= .info(55)
	End If
	If (QQ<>"" And not isnumeric(QQ)) Or (IsQQpic="1" and QQ="") then .GoToErr(46)
	If Len(Honor)>16 or Len(Mail)>50 or Len(HeadPic)>220 or Len(QQ)>20 or Len(Home)>250 Then .GoToErr(47)
	If Not isnumeric(PicW) or Not isnumeric(PicH) Then .GoToErr(48)
	If Int(PicW)>int(.info(56)) or Int(PicH)>int(.info(56)) then
		PicW=.info(54)
		PicH=.info(55)
	End If
	Birthday=Replace(Birthday,",","-")
	If Not isdate(Birthday) then
		Birthday="Birthday=Null"
	Else
		.Cache.clean("Birthday")
		Birthday="Birthday='"&Birthday&"'"
	End If
		.execute("update [User] set "&Birthday&",Sex="&Sex&",PicW="&PicW&",PicH="&PicH&",Mail='"&Mail&"',QQ='"&QQ&"',Honor='"&Honor&"',Pic='"&PicUrl&"',Home='"&Home&"',Sign='"&Sign&"',IsQQpic="&IsQQpic&" where ID="&BBS.MyID)
	Content="<div style='margin:15px;line-height: 150%'><li>资料修改成功！<li><a href=userinfo.asp>返回我的用户控制面版</a><li><a href=index.asp>返回首页</a></div>"
	.ShowTable"修改成功",Content
	Session(CacheName & "Myinfo") = Empty
	End with
End Sub

Sub MyPassword
	Dim S
	BBS.Head "","","修改密码"
	MyManager()
	S="<form style='margin:0' method='POST' action='?action=savemypassword' name='form'>"
	S=S&BBS.Row("<b>旧密码确认：</b><br />请输入旧密码进入确认","<input type='password' name='Password' size='30' maxlength='20' class='text' />","65%","44px")
	S=S&BBS.Row("<b>新的密码(最多14位)：</b><br />请使用除“'”和“|”以及中文以外的字符","<input type='password' name='NewPassword' size='30' maxlength='20' class='text' />","65%","44px")
	S=S&BBS.Row("<b>重复密码：</b><br />请再输一遍确认","<input type='password' name='RePassword' size='30' maxlength='20' class='text' />","65%","44px")
	S=S&BBS.Row("<b>密码问题</b>：<br />忘记密码的提示问题","<input type='text' class='text' name='clue' size=30  maxlength='60' /> 如不改请不要填写","65%","40px")
	S=S&BBS.Row("<b>问题答案</b>：<br />忘记密码的提示问题答案，用于取回论坛密码","<input type='text' class='text' name='answer' size=30  maxlength='60' /> 同上","65%","40px")
	S=S&"<div style=""padding:5px;BACKGROUND: "&BBS.SkinsPIC(2)&";""  align='center'><input type='submit' value='确定修改！' class='button' />&nbsp;&nbsp; <input type='reset' class='button' value='取消重置！'  /></div></form>"
	BBS.ShowTable "修改密码",S
End Sub

Sub savemypassword
	Dim Password,NewPassword,RePassword,Caption,Content,Clue,Answer
	BBS.CheckMake'禁止外部提交
	BBS.Head"","","修改密码"
	MyManager()
	Password=BBS.Fun.GetStr("Password")
	NewPassword=BBS.Fun.GetStr("NewPassword")
	RePassword=BBS.Fun.GetStr("RePassword")
	Clue=BBS.Fun.GetStr("clue")
	Answer=BBS.Fun.GetStr("answer")
	If Password="" or Repassword="" or NewPassword="" Then BBS.GoToErr(36)	
	If Repassword<>NewPassword Then BBS.GoToErr(41)
	If Not BBS.Fun.CheckPassword(Password) Or Not BBS.Fun.CheckPassword(NewPassword) Then BBS.GoToErr(37)
	If BBS.Fun.StrLength(NewPassword)>14 Then BBS.GoToErr(38)
	If md5(Password)<>BBS.MyPassword Then BBS.GoToErr(56)
	IF instr("|12345|123456|1234567|12345678|123456789|1234567890|0123456789|111111|222222|333333|888888|aaaaaa|","|"& Password &"|")>0 or len(Password)<5 Then BBS.GoToErr(40)
	If Clue<>"" or Answer<>"" Then
		If Len(Clue)<3  or Len(Answer)<3 Then BBS.GoToErr(43)
		If not BBS.Fun.CheckIn(Clue) or not BBS.Fun.CheckIn(Answer) Then BBS.GoToErr(44)
		BBS.execute("update [user] set [Clue]='"&Clue&"',Answer='"&MD5(Answer)&"' where ID="&BBS.MyID)
	End IF
	NewPassword=Md5(Newpassword)
	BBS.execute("update [user] set [password]='"&Newpassword&"' where ID="&BBS.MyID)
	BBS.LetMemor "","MyPassword",NewPassword
	Session(CacheName & "Myinfo") = Empty
	Content="<div style='margin:15px;line-height: 150%'><li>密码修改成功!</li><li><a href=userinfo.asp>返回用户控制面版</a></li><li><a href=index.asp>返回首页</a></li></div>"
	BBS.ShowTable "修改成功",Content
End Sub

Sub ForgetPassword
	Dim UserName,rs,S
	BBS.Head"","","找回密码"
	UserName=BBS.Fun.GetStr("UserName")
	If UserName="" Then
	S="<form style='margin:0' method='POST'>"
	S=S&BBS.Row("<b>您的用户名：</b>","<input type='text' name='UserName' size='30' maxlength='20' class='text' /> 请输入注册的用户名称进入确认","90%","22px")
	S=S&"<div style=""padding:5px;BACKGROUND: "&BBS.SkinsPIC(2)&";""  align='center'><input type='submit' value='确定！' class='button' /></div></form>"
	Else
	BBS.CheckMake'禁止外部提交
	If BBS.SafeBuckler(UserName,BBS.MyIP,0) Then BBS.Alert"BBS安全盾已启动！\n\n对不起，你尝试找回密码错误超过3次，今天不能再找回密码了。\n你的信息已被系统记录！","Index.asp"
	set rs=BBS.Execute("select clue from [User] where name='"&UserName&"'")
	S="<form style='margin:0' method='POST' action='?action=getpassword'>"
	S=S&BBS.Row("<b>用户名称：</b>","<input name='UserName' type='hidden' value='"&UserName&"'><b>"& UserName &"</b>","90%","22px")
	S=S&BBS.Row("<b>密码问题：</b>", ""&Rs("clue")&"","90%","22px")
	S=S&BBS.Row("<b>问题答案：</b>","<input type='text' name='Answer' size='30' maxlength='20' class='text' /> 请输入您在注册时填写的问题答案","90%","22px")
	S=S&"<div style=""padding:5px;BACKGROUND: "&BBS.SkinsPIC(2)&";""  align='center'><input type='submit' value='确定！' class='button' />&nbsp;&nbsp; <input type='reset' class='button' value='重置！'  /></div></form>"
	End If
	BBS.ShowTable"找回密码",S
End Sub

Sub GetPassword
	Dim UserName,Clue,Answer,NewPassword,Content
	BBS.Head"","","找回密码"
	UserName=BBS.Fun.GetStr("UserName")
	If BBS.SafeBuckler(UserName,BBS.MyIP,0) Then BBS.Alert"BBS安全盾已启动！\n\n对不起，你尝试找回密码错误超过3次，今天不能再找回密码了。\n你的信息已被系统记录！","Index.asp"
	Answer=BBS.Fun.GetStr("Answer")
	If UserName="" or Answer="" Then BBS.GoToErr(36)
	If Not BBS.Fun.CheckName(UserName) Then BBS.GoToErr(37)
	If not BBS.Fun.CheckIn(Answer) Then BBS.GoToErr(44)
	IF BBS.execute("select name from [User] where name='"&UserName&"' And Answer='"&Md5(Answer)&"'").eof  Then
		BBS.SafeBuckler UserName,BBS.MyIP,1
		BBS.GoToErr(57)
	Else
		Randomize
		NewPassword=int(900000*rnd)+100000
		BBS.execute("update [user] set [password]='"&Md5(NewPassword)&"' where name='"&UserName&"'")
		BBS.execute("update [Admin] set [password]='"&Md5(NewPassword)&"' where name='"&UserName&"'")	
		Content="<div style='margin:15px;line-height: 150%'><li>您成功的通过密码保护的检验！</li><li>用户名称：<font color=red>"&UserName&"</font> &nbsp; 获得新密码：<font color=red>"&NewPassword&"</font></li><li>先记住新密码，请您马上登陆论坛，尽快修改密码！</li></div>"
		BBS.ShowTable "成功通过验证",Content
	End If	
End Sub
%>
