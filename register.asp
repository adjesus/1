<!--#include file="Inc.asp"-->
<!--#include file="inc/md5.asp"-->
<%
Dim Action,Page_Url
If Session(CacheName&"RegTime")+BBS.Info(9)/1440>now() then BBS.GotoErr(24)
If BBS.Info(40)="0" then BBS.GoToErr(23)
Action=request.querystring("action")
If Action <> "" Then
  Page_Url = "?action="&Action
Else
  Page_Url = ""
End If
BBS.Head"Register.asp"&Page_Url,"","注册新用户"

If Len(Action)>10 Then BBS.GoToErr(1)
Select Case Action
Case"agree"
	Register()
Case"check"
	RegSaveData()
Case Else
	RegMain()
End Select
BBS.Footer()
Set BBS =Nothing
Sub RegMain()
	Dim Caption,Content
	Caption="注册协议"
	Content="<div style=""text-align:center""><iframe style=""font-size:12px;width:96%;height:400px;border:#999 1px solid"" frameborder=""0"" src=""inc/agreement.html"" scrolling=""auto"" ></iframe></div>"&_
	"<form method=POST action='?action=agree'><center><input type='submit' class='button' value='同意协议'> <input type='button' class='button' value='我不同意' onClick=window.location.href='index.asp'></center></form>"
	BBS.ShowTable Caption,Content
End Sub

Sub RegSaveData()
With BBS
	.CheckMake'禁止外部提交
	Dim S,Caption,Content,Rs,Name,password,RePassword,Clue,Answer,Mail,PicUrl,headpicoption,PicW,PicH,Home,Sign,QQ,IsQQpic,Sex,Birthday,iCode,UserID,IsDel
	Name=.Fun.GetStr("name")
	password=.Fun.GetStr("password")
	RePassword=.Fun.GetStr("repassword")
	Clue=.Fun.GetStr("clue")
	Answer=.Fun.GetStr("answer")
	Mail=.Fun.GetStr("mail")
	iCode=.Fun.GetStr("iCode")
	If name="" or Password="" or RePassword="" or Mail="" or Clue="" or Answer="" Then .GoToErr(36)	
	If .Fun.StrLength(name)>14 or .Fun.StrLength(name)<2 or .Fun.strLength(password)>14 Then .GoToErr(38)
	If Not .Fun.CheckName(name) OR Not .Fun.CheckPassword(Password) Then .GoToErr(37)
	If instr(lcase(.Info(52)),lcase(Name))>0 Then .GoToErr(37)
	If Not .Execute("SELECT name FROM [user] where Name='"&Name&"'").Eof Then .GoToErr(39)
	IF instr("123456|1234567|12345678|123456789|1111111|222222|333333|888888|aaaaaaa","|"& Password &"|")>0 or len(Password)<6 Then .GoToErr(40)
	If Repassword<>Password Then .GoToErr(41)
	If .Info(13)="1" Then
		If iCode<>Session("iCode") or Session("iCode")="" Then .GotoErr(8)
	End If
	Mail=server.HTMLEnCode(Mail)
	If Not .Fun.IsValidEmail(Mail) Then .GoToErr(42)
	'只允一个邮箱
	If .Info(42)="1" Then
		If Not .Execute("SELECT ID FROM [user] where Mail='"&Mail&"'").Eof Then .GoToErr(49)
	End If
	If .Fun.GetStr("rnd")<>"bd04c9fea4c8" Then .GoToErr(2)
	If Len(Clue)<3  or Len(Answer)<3 Then .GoToErr(43)
	If not .Fun.CheckIn(Clue) or not .Fun.CheckIn(Answer) Then .GoToErr(44)
	PicUrl=lcase(.Fun.HtmlCode(.Fun.GetStr("PicUrl")))
	headpicoption=.Fun.HtmlCode(.Fun.GetStr("headpicoption"))
	If Not .Fun.isInteger(headpicoption) Or Not .Fun.IsUrl(PicUrl) Then .GoToErr(81)
	Home=.Fun.HtmlCode(.Fun.GetStr("Home"))
	Sex=.Fun.GetStr("Sex")
	Birthday=.Fun.GetStr("Birthday")
	QQ=.Fun.GetStr("QQ")
	IsQQpic=.Fun.GetStr("IsQQpic")
	If Instr(Home,"://")=0 Then Home=.Info(1)
	If IsQQpic<>"1" Then IsQQpic="0"
	Sign=Replace(Left(.Fun.Replacehtml(.Fun.GetStr("Sign")),255),"{帖子内容}","")
	PicH=.Fun.GetStr("PicH")
	PicW=.Fun.Getstr("PicW")
	If .Info(57)="1" And (Instr(PicUrl,"://")>0  Or Instr(Lcase(Picurl),"www")>0 Or Instr(Lcase(PicUrl),"..")>0) Then .GotoErr(45)'禁止外部图片
	If PicUrl="" then
		PicUrl="Pic/headpic/"& headpicoption &".gif"
		PicW= .Info(54)
		PicH= .Info(55)
	End If
	If (QQ<>"" And not isnumeric(QQ)) Or (IsQQpic="1" and QQ="") then .GoToErr(46)
	If Len(Clue)>70 Or Len(Answer)>70 or Len(Mail)>50 or Len(PicUrl)>220 or Len(QQ)>20 or Len(Home)>250 Then .GoToErr(47)
	If Not isnumeric(PicW) or Not isnumeric(PicH) Then .GoToErr(48)
	If Int(PicW)>int(.Info(56)) or Int(PicH)>int(.Info(56)) then
		PicW=.Info(54)
		PicH=.Info(55)
	End If
	If Not isdate(Birthday) then
		Birthday="Null"
	Else
		.Cache.clean("Birthday")
		Birthday="'"&Birthday&"'"
	End If
	If .Info(41)="1" then
	IsDel=2'注册核审
	S="<li>你的注册信息已提交，请您等待管理员的审核！</li>"
	Else
	IsDel=0
	S="<li>现在<a href='login.asp'>登陆</a>？</li>"
	End If
	.Execute("Insert into [User](Name,[Password],Clue,Answer,Mail,Home,Sex,IsQQpic,Birthday,QQ,Pic,PicW,PicH,RegTime,LastTime,Sign,Regip,Coin,Isdel,GoodNum,EssayNum,Mark,BankSave,isShow,isVip,isSign,LoginNum,GameCoin,BankTime)VALUES('"&Name&"','"&Md5(password)&"','"&Clue&"','"&Md5(Answer)&"','"&Mail&"','"&Home&"',"&Sex&","&IsQQpic&","&Birthday&",'"&QQ&"','"&PicUrl&"',"&PicW&","&PicH&",'"&.NowBbsTime&"','"&.NowBbsTime&"','"&Sign&"','"&.MyIP&"',100,"&IsDel&",0,0,0,0,0,0,0,0,0,'"&.NowBbsTime&"')")
	UserID=.Execute("Select ID From[User] where Name='"&Name&"'")(0)
	.UpdateGrade UserID,0,0
	.Execute("update [Config] set NewUser='"&name&"',UserNum=UserNum+1")
'自动发送留言
	If .Info(43)="1" Then
		.Execute("insert into [sms](name,MyName,Content,ubbString) values('自动送信系统','"&name&"','"&.Info(46)&"',',')")
		.Execute("update [User] set NewSmsNum=1,SmsSize=1 Where Name='"&name&"'")
	End If
	Caption="注册成功"
	Content="<div style='margin:15px;line-height: 150%'><b>恭喜您成为本论坛会员</b>"&S&"<li>返回<a href='index.asp'>首页</a></li></div>"
	.ShowTable Caption,Content
	Session(CacheName&"RegTime")=Now()
	S=Replace(Join(.InfoUpdate,","),","&.InfoUpdate(5)&","&.InfoUpdate(6)&",",","&Int(.InfoUpdate(5))+1&","&Name&",")
	.Cache.Add "InfoUpdate",S,dateadd("n",2000,.NowBBSTime)
	End with
End Sub
Sub Register()
	Dim S
	S="<form style='margin:0' method='POST' action='?action=check' name='form'>"
	S=S&"<div style='text-align:left;border-bottom:1px solid "&BBS.SkinsPIC(0)&";padding:3px;font-weight:bold;BACKGROUND: "&BBS.SkinsPIC(2)&";'><img src='Images/icon/inn.gif' align='absmiddle' atl='必填资料' /> 必填资料</div>"
	S=S&BBS.Row("<b>用户名称：</b><br />注册用户名不能超过14个字符（7个汉字）","<input type='text' class='text' id='name' name='name' maxlength='20' /> <input onClick='Check()' type='button' value='检测用户名' />","65%","40px")
	S=S&BBS.Row("<b>密码(最少6位,最多16位)</b>：<br>请使用除“'”和“|”和中文以外的字符","<input type='password' name='password' id='password' maxlength='20' onkeyup='javascript:SetPwdStrengthEx(document.forms[0],this.value);' /><br /><div style='text-align:center;position:absolute;line-height:18px;background-color:#EBEBEB;border-bottom:solid 1px #BEBEBE;'><div id='idSM1' style='height:18px;float:left;width:50px;border-right:solid 1px #BEBEBE;'><span id='idSMT1' style='display:none;'>弱</span></div><div id='idSM2' style='height:18px;float:left;width:60px;border-right:solid 1px #BEBEBE;border-left:solid 1px #fff'><span id='idSMT0' style='color:#666'>未能评级</span><span id='idSMT2' style='display:none;'>中</span></div><div id='idSM3' style='text-align:center;height:18px;float:left;width:60px;border-left:solid 1px #fff;border-right:solid 1px #BEBEBE;'><span id='idSMT3' style='display:none;'>强</span></div></div>","65%","43px")
	S=S&BBS.Row("<b>重复密码</b>：<br />请再输一遍确认","<input type='password' name='repassword' maxlength='20'>","65%","40px")
	If BBS.Info(13)="1" Then S=S&BBS.Row("<b>请输入右边的验证码：</b>",BBS.GetiCode,"65%","")
	S=S&BBS.Row("<b>您的性别：</b>","<input name='sex' type='radio' value='1' checked class=checkbox /><img src='Images/icon/male.gif' align='absmiddle' atl='帅哥' /> 帅哥&nbsp;&nbsp;<input type='radio' name='sex' value='0' class=checkbox /><img src='Images/icon/female.gif' align='absmiddle' atl='靓女' /> 靓女","65%","")
	S=S&BBS.Row("<b>密码问题</b>：<br />忘记密码的提示问题","<input type='text' class='text' name='clue' size=30  maxlength='60' />","65%","40px")
	S=S&BBS.Row("<b>问题答案</b>：<br />忘记密码的提示问题答案，用于取回论坛密码","<input type='text' class='text' name='answer' size=30  maxlength='60' />","65%","40px")
	S=S&BBS.Row("<b>OICQ号码：</b><br />填写您的QQ地址，方便与他人的联系","<input type='text' class='text' name='QQ'  maxlength='15'> <input type='checkbox' onclick='QQpic()' name='isqqpic' id='isqqpic' value='1' class=checkbox>启用QQ形象作为头像","65%","40px")
	S=S&BBS.Row("<b>Email地址</b>：<br>请输入有效的邮件地址","<input type='text' class='text' name='mail' size='30' maxlength='30'>","65%","40px")
	S=S&"<div style='text-align:left;border-bottom:1px solid "&BBS.SkinsPIC(0)&";padding:3px;font-weight:bold;BACKGROUND: "&BBS.SkinsPIC(2)&";'><img src='Images/icon/inn.gif' align='absmiddle' atl='选填资料' /> 选填资料</div>"
	S=S&"<div id='Rowhow' display='none'>"
	S=S&BBS.Row("<b>生日：</b>","<input  type='text' class='text' name='birthday' id='birthday' size='10' maxlength='10' readonly='true' onfocus=""show_cele_date('birthday','','',this)"" />","65%","20px")
	S=S&BBS.Row("<b>主页：</b><br />填写你的个人主页，让大家见识见识！","<input type='text' class='text' name='home' size='30' maxlength='200' />","65%","40px")
	S=S&"<div id='showpic'>"
	S=S&BBS.Row("<b>选择论坛头像：</b><br />使用论坛自带的图像",HeadPicOpt() &"<img src='Pic/headpic/1.gif' id='pic' name='pic' /> <input onclick=""openwin('preview.asp?Action=HeadPic',680,400,'yes')"" type=button value='全部头像' class='button'  />","65%","")
	S=S&BBS.Row("<b>自定义头像：</b><br />如果图像位置中有连接图片将以自定义的为主","<input id='picurl' name='picurl' size='40' maxlength='100' /> 完整Url地址<br />图像宽度：<input type='text' class='text' name='picw' size='6' value='"& BBS.Info(54) &"' /> 高度：<input type=text name='pich' size='6' value='"&BBS.Info(55)&"' />(最大限度:120)","65%","")
	S=S&"<input type=""hidden"" name=""rnd"" value=""bd04c9fea4c8"" /></div>"
	S=S&BBS.Row("<b>个性签名：</b><br />文字将出现在您发表的文章的结尾处<br />体现您的个性(最多255个字符)","<TEXTAREA name='sign' rows='4' style='width: 98%;'></textarea>","65%","")
	S=S&"</div><div style=""padding:5px;BACKGROUND: "&BBS.SkinsPIC(2)&";""  align='center'><input type='submit' value='确定注册！' class='button' />&nbsp;&nbsp; <input type='reset' value='取消重写！' class='button' /></div></form>"
	BBS.ShowTable"新用户注册",S
	%>
<script language="JavaScript" type="text/javascript" src="Inc/pswdplc.js"></script>
<script language="JavaScript" type="text/javascript" src="Inc/InputDate.js"></script>
<script language="JavaScript" type="text/javascript">
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
}
function Check(){
openwin("preview.asp?Action=CheckName&name="+document.getElementById("name").value,300,30,"no");
}
</script>
<%
End Sub

Function HeadPicOpt()
	Dim S,i
	for i=2 to Int(BBS.Info(53))
		S=S&"<option value='"&i&"'>"&i&"</option>"
	Next
	HeadPicOpt="<select name='headpicoption' id='headpicoption' onChange='ShowPic()'><option value='1' selected>1</option>"&S&"</select>"
End Function
%>
