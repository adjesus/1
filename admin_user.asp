<!--#include file="Admin_Check.asp"-->
<!--#include file="Inc/Md5.asp"-->
<script language="JavaScript" type="text/javascript" src="Inc/pswdplc.js"></script>
<%
Head()
Select Case Lcase(request.querystring("Action"))
Case"edituser"
	CheckString "21"
	EditUser
Case"saveuser"
	CheckString "21"
	SaveUser
Case"adminok"
	AdminOK
Case"executesql"
	CheckString "56"
	ExecuteSql
Case"sqlpassword"
	CheckString "56"
	SqlPassword
End select
Footer()

Sub SqlPassword
	Dim NewP,OldP
	NewP=BBS.Fun.GetStr("new")
	oldP=BBS.Fun.GetStr("old")
	If NewP="" And oldP="" then 
		Showtable"更改执行SQL语句的密码","<form method='post'>原密码：<input type='password' class='text' name='old' value='' size='20' /> 新密码：<input type='text' class='text' name='new' ><input type='submit' class='button' value='确 定' /></Form>"
	Else
		If newP="" or oldP="" then goback"","":Exit Sub
		If BBS.Execute("Select * From [Config] where SqlPassword='"&MD5(oldP)&"'").Eof Then Goback"","修改失败，原密码不正确！":Exit Sub
		BBS.Execute("update [Config] Set SqlPassword='"&MD5(newP)&"'")
		BBS.NetLog"操作后台_更改执行SQL语句的密码"
		Suc "","成功更改执行SQL语句的密码!","?Action=ExecuteSql"
	End If
End Sub

Sub ExecuteSql
	Dim Sql,Password,Caption,Content,S
	Sql=Request.Form("sql")
	Password=BBS.Fun.GetStr("password")
	Caption="执行SQL语句"
	Content="<form method='post'>密码：<input type='password' class='text' name='password' value='' size='20' /> <a href='?action=SqlPassword'>更改密码</a><br />指令：<input type='text' class='text' name='sql' value='"&replace(Sql,"'","&#39;")&"' style='width:90%'><br>注意：此操作不可恢复，如果对SQL语法不了解，请慎用！<input type='button' class='button' onclick=""if(confirm('注意！操作不当有可能破坏数据库！\n\n您确定要执行SQL语句吗？'))form.submit()"" value='确定执行' /></Form>"
	ShowTable Caption,Content
	If Sql<>"" then
		If Password="" Then Goback"","":Exit Sub
		If BBS.Execute("Select * From [Config] where SqlPassword='"&MD5(Password)&"'").Eof Then Goback"","密码错误":Exit Sub
		On Error Resume Next 
		BBS.Execute(Sql)
		If err.number=0 then
			Caption="执行成功":Content="<li>Sql语句正确，已经成功的执行了下面这条语句！<li><font color=red>"&Sql&"</font></li>"
			BBS.NetLog"操作后台_成功执行SQL语句：<br>"&Sql&""
		Else
			Caption="错误信息":Content="<li>不能执行，语句有问题，具体出错如下：</li><li>"&Err.Description&"</li>"
			Err.clear
		End if
		ShowTable Caption,Content
	End if
End Sub


Sub AdminOK
Dim ID,S,isOK
Dim Menu(5,10),I,J,Strings,Name,Password,Temp
ID=Replace(Request.Form("ID")," ","")
Name=Replace(Request("name"),"'","")
IsOK=true
If Instr(AdminString,",22,")=0 Then
	IsOK=False
	If lcase(BBS.MyName)<>lcase(Name) Then Goback"","你没有编辑其它管理员的权限！":Exit Sub
End If
Password=Request("password")
If ID<>"" or Password<>"" Then
	If Password<>"" Then 
		If len(Password)<6 then goback"","后台密码不能设得太简单！为了安全，建议用大小写字母加数字而且不要少于8位的密码！":Exit sub
		Password=MD5(Password)
		Temp="[Password]='"&Password&"'"
		If lcase(Name)=Lcase(BBS.GetMemor("Admin","AdminName")) Then
			BBS.LetMemor "Admin","AdminName",Name
			BBS.LetMemor "Admin","AdminPassword",Password
		End If	
	End If
	
	If Temp<>"" Then
		If ID<>"" Then
			ID=","&ID&","
			Temp=Temp&",Strings='"&ID&"'"
		End If
	ELse
		ID=","&ID&","
		Temp="Strings='"&ID&"'"
	End IF
	BBS.execute("update [Admin] Set "&Temp&" where Name='"&Name&"' And BoardID=0")
	S="更改管理员："&Name&" 的后台权限成功"
	BBS.NetLog"操作后台_"&S
	Suc "",S,"Admin_Action.asp?Action=TopAdmin"
Else
Menu(0,0)="系统设置"
Menu(0,1)="论坛信息设置"
Menu(0,2)="论坛统计设置"
Menu(0,3)="公告发布管理"
Menu(0,4)="帖间广告管理"
Menu(0,5)="论坛联盟管理"
Menu(0,6)="I P 封锁管理"
Menu(0,7)="论坛日志系统"
Menu(0,8)="更新论坛缓存"
Menu(0,9)="论坛调用系统"
Menu(1,0)="论坛版块"         
Menu(1,1)="论坛版面管理"
Menu(1,2)="添加论坛分类"       
Menu(1,3)="添加论坛版面"
Menu(2,0)="用户管理"
Menu(2,1)="用户批量管理"
Menu(2,2)="设置管理人员"
Menu(2,3)="设置论坛版主"
Menu(2,4)="设置 VIP用户"
Menu(2,5)="恢复删除用户"
Menu(2,6)="设置特别等级"
Menu(2,7)="用户等级管理"
Menu(3,0)="帖子留言"         
Menu(3,1)="批量删除帖子"         
Menu(3,2)="批量移动帖子"        
Menu(3,3)="批量删除留言"
Menu(3,4)="群发信件留言"         
Menu(3,5)="上传文件管理"
Menu(3,6)="论坛回收站"
Menu(4,0)="论坛DIY"  
Menu(4,1)="论坛菜单管理"
Menu(4,2)="修改注册协议"
Menu(4,3)="风格模板管理"
Menu(4,4)="论坛银行管理"
Menu(4,5)="论坛帮派管理"
Menu(5,0)="论坛数据"
Menu(5,1)="压缩数据库"        
Menu(5,2)="备份数据库"        
Menu(5,3)="恢复数据库"    
Menu(5,4)="数据表管理"
Menu(5,5)="论坛整理修复"     
Menu(5,6)="执行SQL语句"
Menu(5,7)="空间占用情况"
Menu(5,8)="服务器检测"
If Name="" Then goback"","":exit Sub
Set Rs=BBS.Execute("Select Strings from [Admin] where name='"&Name&"' and boardID=0")
If Rs.Eof Then Goback"","数据不存在":Exit Sub
Strings=Rs(0)
Rs.Close
Response.Write"<form method='post' style='margin:0px' action='?Action=AdminOK&Name="&Name&"'>"
Response.Write"<div class='mian'><div class='top'>管理员 "&Name&" 后台权限设置</div>"
Response.Write"<div class='divtr2' style='padding:4px;'><strong>后台密码：</strong><input type='text' name='password' size='20' class='text' onkeyup='javascript:SetPwdStrengthEx(document.forms[0],this.value);' /> <div style='text-align:center;position:absolute;line-height:18px;background-color:#EBEBEB;border-bottom:solid 1px #BEBEBE;'><div id='idSM1' style='height:18px;float:left;width:50px;border-right:solid 1px #BEBEBE;'><span id='idSMT1' style='display:none;'>弱</span></div><div id='idSM2' style='height:18px;float:left;width:60px;border-right:solid 1px #BEBEBE;border-left:solid 1px #fff'><span id='idSMT0' style='color:#666'>未能评级</span><span id='idSMT2' style='display:none;'>中</span></div><div id='idSM3' style='text-align:center;height:18px;float:left;width:60px;border-left:solid 1px #fff;border-right:solid 1px #BEBEBE;'><span id='idSMT3' style='display:none;'>强</span></div></div>"
Response.Write"<br>密码如果不改请不要填。(为了安全，建议设置强度复杂的密码。)</div><div class='divtr1' style='padding:3px;'>"

for i=0 to ubound(menu,1)
Response.Write"<div style='padding:3px;'><b>"&menu(i,0)&"</b><br>"
for j=1 to ubound(menu,2)
If isempty(menu(i,j)) then exit for
Response.Write" "&i&j&"<input type='checkbox' name='ID' value='"&i&j&"' "
if instr(Strings,","&i&j&",")<>0 then response.write "checked"
If not IsOk Then Response.write " disabled='true'"
Response.Write" />"&Menu(i,j)
If j mod 5 =0 Then Response.write "<br>"
next
Response.Write"</div>"
next
Response.Write"</div><div class='bottom'><input type='checkbox' name='chkall' value='on' onClick='CheckAll(this.form)'" 
If not IsOk Then Response.write " disabled='true'"
Response.Write" />全选&nbsp;　<input type='submit' class='button' value='提 交'><input class='button' type='reset' value='重 置'></div></div></form>"
End IF
End Sub



Sub EditUser
	Dim ID,Temp,Rs,UserName,Sqlwhere
	ID=Request("ID")
	UserName=Request("Name")
	Sqlwhere="ID="&ID
	If UserName<>"" then Sqlwhere="Name='"&UserName&"'"
	Response.Write "<form method='post' style='margin:0px' action='?Action=SaveUser'>"
	Set Rs=BBS.Execute("select Name,Password,Clue,Answer,Sex,Mail,Birthday,Home,QQ,isQQpic,Pic,PicW,PicH,Sign,Regtime,RegIp,Lasttime,LastIp,EssayNum,GoodNum,Mark,Coin,BankSave,GameCoin,Honor,Faction,LoginNum,isDel,isVip,isShow,isSign,GradeID,GradeFlag,NewSmsNum,BankTime,ID From [USER] where "&Sqlwhere)
	If Rs.eof Then
		Goback "","该用户不存在！"
		Exit Sub
	End If
	Response.Write"<div class='mian'><div class='top'><a style='FLOAT: right;color:#FFF' href='Admin_ActionList.asp?action=UserList'>返回用户管理&nbsp;</a>修改用户： "&Rs(0)&"</div>"
	Response.Write"<div class='divtr2' style='padding:3px;line-height:20px'><div style='FLOAT: right;width:50%'><fieldset><legend>快捷操作</legend><a href='Admin_ActionList.asp?action=setgrade&Name="&Rs(0)&"'>设置特别等级组</a><br><a href='admin_action.asp?action=BoardAdmin&Name="&Rs(0)&"'>提升版主</a><br><a href=#this onclick=""checkclick('删除该用户（包括其帖子）\n\n删除后将不能恢复！您确定要删除吗？','Admin_Confirm.asp?Action=UpdateUserList&point=2&ID="&Rs(35)&"')"">完全删除</a><br><a href='Admin_Action.asp?action=AddLockIp&IP="&Rs(15)&"&Readme=封("&Rs(0)&")的注册IP'>封锁注册IP</a><br><a href='Admin_Action.asp?action=AddLockIp&IP="&Rs(17)&"&Readme=封("&Rs(0)&")的最后登陆IP'> 封锁最后登陆IP</a></fieldset></div><fieldset><legend>用户信息</legend>所在等级组："&BBS.GetGradeName(Rs(31),Rs(32))&"<br>注册会员时间："&Rs(14)&" <br>最后登陆时间："&Rs(16)&" <br>注册会员时IP记录："&Rs(15)&"<br>最后登陆时IP记录："&Rs(17)&"</fieldset></div>"
	Response.Write"<div style='text-align:left; padding:3px' class='divth'><b>用户注册信息</b></div>"
	DIVTR"用户名称：","","<input name='ID' type='hidden' value='"&Rs(35)&"'><input type='text' class='text' name='Name' size='20' value='"&Rs(0)&"' />",25,1
	DIVTR"用户密码：","","<input type='text' class='text' name='Password' size='20' value='' /> 不改请不要填",25,2
	DIVTR"密码问题：","","<input type='text' class='text' name='Clue' size='20' value='"&Rs(2)&"' />",25,1
	DIVTR"密码答案：","","<input type='text' class='text' name='Answer' size='20' value='' /> 不改请不要填",25,2
	DIVTR"性别：","",GetRadio("Sex","女",Rs(4),0)&GetRadio("Sex","男",Rs(4),1),25,1
	DIVTR"邮箱：","","<input type='text' class='text' name='Mail' size='20' value='"&Rs(5)&"' />",25,2
	DIVTR"生日：","","<input type='text' class='text' name='Birthday' size='20' value='"&Rs(6)&"' />",25,1
	DIVTR"主页：","","<input type='text' class='text' name='Home' size='40' value='"&Rs(7)&"' />",25,2
	DIVTR"QQ号码：","","<input type='text' class='text' name='QQ' size='20' value='"&Rs(8)&"' />",25,1
	DIVTR"启用QQ形象作为头像：","",GetRadio("isQQpic","否",Rs(9),0)&GetRadio("isQQpic","是",Rs(9),1)&"(QQ号码必须填写)",25,2
	DIVTR"头像：","","<input type='text' class='text' name='Pic' size='40' value='"&Rs(10)&"' />",25,1
	DIVTR"头像尺寸：","","宽：<input type='text' class='text' name='PicW' size='5' value='"&Rs(11)&"' /> 高：<input type='text' class='text' name='PicH' size='5' value='"&Rs(12)&"' />",25,2
	DIVTR"签名：","","<TEXTAREA name='sign' rows='4' style='width: 98%;'>"&Rs(13)&"</textarea>",60,1
	Response.Write"<div style='text-align:left; padding:3px' class='divth'><b>用户论坛信息</b></div>"
	DIVTR"总帖数：","","<input type='text' class='text' name='EssayNum' size='6' value='"&Rs(18)&"' />",25,1
	DIVTR"精华帖数：","","<input type='text' class='text' name='GoodNum' size='6' value='"&Rs(19)&"' />",25,2
	DIVTR"积分：","","<input type='text' class='text' name='Mark' size='6' value='"&Rs(20)&"' />",25,1
	DIVTR"金钱：","","<input type='text' class='text' name='Coin' size='6' value='"&Rs(21)&"' />",25,2
	DIVTR"存款：","","<input type='text' class='text' name='BankSave' size='6' value='"&Rs(22)&"' />",25,1
	DIVTR"游戏币：","","<input type='text' class='text' name='GameCoin' size='6' value='"&Rs(23)&"' />",25,2
	DIVTR"登陆次数：","","<input type='text' class='text' name='LoginNum' size='6' value='"&Rs(26)&"' />",25,1
	DIVTR"头衔：","","<input type='text' class='text' name='Honor' size='30' value='"&Rs(24)&"' />",25,2
	DIVTR"帮派：","","<input type='text' class='text' name='Faction' size='30' value='"&Rs(25)&"' />",25,1
	Response.Write"<div style='text-align:left; padding:3px' class='divth'><b>操作用户选项</b></div>"
	DIVTR"暂时删除：","",GetRadio("isDel","否",Rs(27),0)&GetRadio("isDel","是",Rs(27),1),25,1
	DIVTR"VIP会员：","",GetRadio("isVip","否",Rs(28),0)&GetRadio("isVip","是",Rs(28),1),25,2
	DIVTR"屏蔽帖子：","",GetRadio("isShow","否",Rs(29),0)&GetRadio("isShow","是",Rs(29),1),25,1
	DIVTR"屏蔽签名：","",GetRadio("isSign","否",Rs(30),0)&GetRadio("isSign","是",Rs(30),1),25,2
	Response.Write"<div class='bottom'><input type='submit' class='button' value='提 交'><input class='button' type='reset' value='重 置'></div></div></form>"
	Rs.Close
End Sub

Sub SaveUser
Dim OldName,ID,AllTable,i,Temp
Dim	Name,Password,Clue,Answer,Sex,Mail,Birthday,Home,QQ,isQQpic,Pic,PicW,PicH,Sign,EssayNum,GoodNum,Mark,Coin,BankSave,GameCoin,Honor,Faction,LoginNum,isDel,isVip,isShow,isSign,GradeFlag,GradeID
	ID=Request.Form("ID")
	Name=Replace(BBS.Fun.Getform("Name"),"'","")
	If Name="" Then GoBack"","":Exit Sub
	Password=BBS.Fun.GetForm("Password")
	Clue=BBS.Fun.GetForm("Clue")
	Answer=BBS.Fun.GetForm("Answer")
	Sex=BBS.Fun.GetForm("Sex")
	Mail=BBS.Fun.GetForm("Mail")
	Birthday=BBS.Fun.GetForm("Birthday")
	Home=BBS.Fun.GetForm("Home")
	QQ=BBS.Fun.GetForm("QQ")
	isQQpic=BBS.Fun.GetForm("isQQpic")
	Pic=BBS.Fun.GetForm("Pic")
	PicW=Request.Form("PicW")
	PicH=Request.Form("PicH")
	Sign=BBS.Fun.GetForm("Sign")
	EssayNum=Request.Form("EssayNum")
	GoodNum=Request.Form("GoodNum")
	Mark=Request.Form("Mark")
	Coin=Request.Form("Coin")
	BankSave=Request.Form("BankSave")
	GameCoin=Request.Form("GameCoin")
	Honor=BBS.Fun.GetForm("Honor")
	Faction=BBS.Fun.GetForm("Faction")
	LoginNum=Request.Form("LoginNum")
	isDel=Request.Form("isDel")
	isVip=Request.Form("isVip")
	isShow=Request.Form("isShow")
	isSign=Request.Form("isSign")
	if not isnumeric(PicW) or not isnumeric(PicH) or not isnumeric(EssayNum) or not isnumeric(GoodNum) or not isnumeric(Mark) or not isnumeric(Coin) or not isnumeric(BankSave) or not isnumeric(GameCoin) or not isnumeric(LoginNum) or not isnumeric(isDel) or not isnumeric(isVip) or not isnumeric(isShow) or not isnumeric(isSign) then
		GoBack"","一些项必需用数字填写":Exit Sub
	End If
	Set Rs=BBS.Execute("select name,GradeID,GradeFlag From[User] where ID="&ID&"")
	If Rs.eof Then
		GoBack"","这个用户根本不存在！":Exit Sub
	Else
		OldName=Rs(0)
		GradeID=Rs(1)
		GradeFlag=Rs(2)
	End If
	Rs.close
	If Password<>"" Then
		If len(password)<6 Then Goback"","密码不能少于6位":Exit sub
		Password="[Password]='"&Md5(Password)&"',"
	End If
	If Answer<>"" Then Answer="Answer='"&Md5(Answer)&"',"
	If Isdate(Birthday) Then Birthday="Birthday='"&Birthday&"'," Else Birthday="Birthday=null,"
	
	If lcase(Name)<>Lcase(OldName) Then
		If Not BBS.Execute("select name From[User] where Name='"&Name&"' And ID<>"&ID&"").eof Then
			GoBack"","新用户名称已经被注册了,不能改名！":Exit Sub
		End If
	End If
	Temp="update [User] Set "&Password&" Clue='"&Clue&"',"&Answer&"Sex="&Sex&",Mail='"&Mail&"',"&Birthday&" Home='"&Home&"',QQ='"&QQ&"',isQQpic="&isQQpic&",Pic='"&Pic&"',PicW="&PicW&",PicH="&PicH&",Sign='"&Sign&"',EssayNum="&EssayNum&",GoodNum="&GoodNum&",Mark="&Mark&",Coin="&Coin&",BankSave="&BankSave&",GameCoin="&GameCoin&",Honor='"&Honor&"',Faction='"&Faction&"',LoginNum="&LoginNum&",isDel="&isDel&",isVip="&isVip&",isShow="&isShow&",isSign="&isSign&" where ID="&ID
	BBS.Execute(Temp)
	OldName=Replace(OldName,"'","''")
	Temp="更改用户“"&OldName&"”的资料!"
	If lcase(Name)<>Lcase(OldName) Then
		AllTable=Split(BBS.BBStable(0),",")
		For i=0 To uBound(AllTable)
			BBS.Execute("Update [Bbs"&AllTable(i)&"] Set Name='"&Name&"' where Name='"&OldName&"'")
		Next
		BBS.Execute("Update [TopicVoteUser] Set [User]='"&Name&"' where [User]='"&OldName&"'")
		BBS.Execute("Update [Topic] Set Name='"&Name&"' where Name='"&OldName&"'")
		BBS.Execute("Update [Sms] Set MyName='"&Name&"' where MyName='"&OldName&"'")
		BBS.Execute("Update [Sms] Set Name='"&Name&"' where Name='"&OldName&"'")
		BBS.Execute("Update [Placard] Set Name='"&Name&"' where Name='"&OldName&"'")
		BBS.Execute("Update [User] Set Name='"&Name&"' where Name='"&OldName&"'")
		BBS.Execute("Update [Admin] Set Name='"&Name&"' where Name='"&OldName&"'")
		Temp=Temp&" 用户名改为“"&Name&"”"
	End If
'======>>>更新等级
	Set Rs=BBS.Execute("select BoardID from [Admin] where name='"&Name&"' order by BoardID")
		IF Not Rs.eof Then
			IF Rs(0)=0 Then
				GradeFlag=9 
			ElseIF Rs(0)=-1 Then
				GradeFlag=8
			Else
				GradeFlag=7
			End if
		End IF
		Rs.Close		
	If IsVIP=1 and GradeFlag=0 Then GradeFlag=4	
	If GradeFlag=1 Then'如果为特殊组
		If BBS.Execute("Select ID From [grade] where ID="&GradeID).Eof Then GradeFlag=0
	End IF
	BBS.UpdateGrade ID,EssayNum,GradeFlag
'<<-----
	BBS.NetLog "操作后台_"&Temp
	Suc"",Temp,"?Action=EditUser&ID="&ID
End Sub


%>