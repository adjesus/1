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
		Showtable"����ִ��SQL��������","<form method='post'>ԭ���룺<input type='password' class='text' name='old' value='' size='20' /> �����룺<input type='text' class='text' name='new' ><input type='submit' class='button' value='ȷ ��' /></Form>"
	Else
		If newP="" or oldP="" then goback"","":Exit Sub
		If BBS.Execute("Select * From [Config] where SqlPassword='"&MD5(oldP)&"'").Eof Then Goback"","�޸�ʧ�ܣ�ԭ���벻��ȷ��":Exit Sub
		BBS.Execute("update [Config] Set SqlPassword='"&MD5(newP)&"'")
		BBS.NetLog"������̨_����ִ��SQL��������"
		Suc "","�ɹ�����ִ��SQL��������!","?Action=ExecuteSql"
	End If
End Sub

Sub ExecuteSql
	Dim Sql,Password,Caption,Content,S
	Sql=Request.Form("sql")
	Password=BBS.Fun.GetStr("password")
	Caption="ִ��SQL���"
	Content="<form method='post'>���룺<input type='password' class='text' name='password' value='' size='20' /> <a href='?action=SqlPassword'>��������</a><br />ָ�<input type='text' class='text' name='sql' value='"&replace(Sql,"'","&#39;")&"' style='width:90%'><br>ע�⣺�˲������ɻָ��������SQL�﷨���˽⣬�����ã�<input type='button' class='button' onclick=""if(confirm('ע�⣡���������п����ƻ����ݿ⣡\n\n��ȷ��Ҫִ��SQL�����'))form.submit()"" value='ȷ��ִ��' /></Form>"
	ShowTable Caption,Content
	If Sql<>"" then
		If Password="" Then Goback"","":Exit Sub
		If BBS.Execute("Select * From [Config] where SqlPassword='"&MD5(Password)&"'").Eof Then Goback"","�������":Exit Sub
		On Error Resume Next 
		BBS.Execute(Sql)
		If err.number=0 then
			Caption="ִ�гɹ�":Content="<li>Sql�����ȷ���Ѿ��ɹ���ִ��������������䣡<li><font color=red>"&Sql&"</font></li>"
			BBS.NetLog"������̨_�ɹ�ִ��SQL��䣺<br>"&Sql&""
		Else
			Caption="������Ϣ":Content="<li>����ִ�У���������⣬����������£�</li><li>"&Err.Description&"</li>"
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
	If lcase(BBS.MyName)<>lcase(Name) Then Goback"","��û�б༭��������Ա��Ȩ�ޣ�":Exit Sub
End If
Password=Request("password")
If ID<>"" or Password<>"" Then
	If Password<>"" Then 
		If len(Password)<6 then goback"","��̨���벻�����̫�򵥣�Ϊ�˰�ȫ�������ô�Сд��ĸ�����ֶ��Ҳ�Ҫ����8λ�����룡":Exit sub
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
	S="���Ĺ���Ա��"&Name&" �ĺ�̨Ȩ�޳ɹ�"
	BBS.NetLog"������̨_"&S
	Suc "",S,"Admin_Action.asp?Action=TopAdmin"
Else
Menu(0,0)="ϵͳ����"
Menu(0,1)="��̳��Ϣ����"
Menu(0,2)="��̳ͳ������"
Menu(0,3)="���淢������"
Menu(0,4)="���������"
Menu(0,5)="��̳���˹���"
Menu(0,6)="I P ��������"
Menu(0,7)="��̳��־ϵͳ"
Menu(0,8)="������̳����"
Menu(0,9)="��̳����ϵͳ"
Menu(1,0)="��̳���"         
Menu(1,1)="��̳�������"
Menu(1,2)="�����̳����"       
Menu(1,3)="�����̳����"
Menu(2,0)="�û�����"
Menu(2,1)="�û���������"
Menu(2,2)="���ù�����Ա"
Menu(2,3)="������̳����"
Menu(2,4)="���� VIP�û�"
Menu(2,5)="�ָ�ɾ���û�"
Menu(2,6)="�����ر�ȼ�"
Menu(2,7)="�û��ȼ�����"
Menu(3,0)="��������"         
Menu(3,1)="����ɾ������"         
Menu(3,2)="�����ƶ�����"        
Menu(3,3)="����ɾ������"
Menu(3,4)="Ⱥ���ż�����"         
Menu(3,5)="�ϴ��ļ�����"
Menu(3,6)="��̳����վ"
Menu(4,0)="��̳DIY"  
Menu(4,1)="��̳�˵�����"
Menu(4,2)="�޸�ע��Э��"
Menu(4,3)="���ģ�����"
Menu(4,4)="��̳���й���"
Menu(4,5)="��̳���ɹ���"
Menu(5,0)="��̳����"
Menu(5,1)="ѹ�����ݿ�"        
Menu(5,2)="�������ݿ�"        
Menu(5,3)="�ָ����ݿ�"    
Menu(5,4)="���ݱ����"
Menu(5,5)="��̳�����޸�"     
Menu(5,6)="ִ��SQL���"
Menu(5,7)="�ռ�ռ�����"
Menu(5,8)="���������"
If Name="" Then goback"","":exit Sub
Set Rs=BBS.Execute("Select Strings from [Admin] where name='"&Name&"' and boardID=0")
If Rs.Eof Then Goback"","���ݲ�����":Exit Sub
Strings=Rs(0)
Rs.Close
Response.Write"<form method='post' style='margin:0px' action='?Action=AdminOK&Name="&Name&"'>"
Response.Write"<div class='mian'><div class='top'>����Ա "&Name&" ��̨Ȩ������</div>"
Response.Write"<div class='divtr2' style='padding:4px;'><strong>��̨���룺</strong><input type='text' name='password' size='20' class='text' onkeyup='javascript:SetPwdStrengthEx(document.forms[0],this.value);' /> <div style='text-align:center;position:absolute;line-height:18px;background-color:#EBEBEB;border-bottom:solid 1px #BEBEBE;'><div id='idSM1' style='height:18px;float:left;width:50px;border-right:solid 1px #BEBEBE;'><span id='idSMT1' style='display:none;'>��</span></div><div id='idSM2' style='height:18px;float:left;width:60px;border-right:solid 1px #BEBEBE;border-left:solid 1px #fff'><span id='idSMT0' style='color:#666'>δ������</span><span id='idSMT2' style='display:none;'>��</span></div><div id='idSM3' style='text-align:center;height:18px;float:left;width:60px;border-left:solid 1px #fff;border-right:solid 1px #BEBEBE;'><span id='idSMT3' style='display:none;'>ǿ</span></div></div>"
Response.Write"<br>������������벻Ҫ�(Ϊ�˰�ȫ����������ǿ�ȸ��ӵ����롣)</div><div class='divtr1' style='padding:3px;'>"

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
Response.Write" />ȫѡ&nbsp;��<input type='submit' class='button' value='�� ��'><input class='button' type='reset' value='�� ��'></div></div></form>"
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
		Goback "","���û������ڣ�"
		Exit Sub
	End If
	Response.Write"<div class='mian'><div class='top'><a style='FLOAT: right;color:#FFF' href='Admin_ActionList.asp?action=UserList'>�����û�����&nbsp;</a>�޸��û��� "&Rs(0)&"</div>"
	Response.Write"<div class='divtr2' style='padding:3px;line-height:20px'><div style='FLOAT: right;width:50%'><fieldset><legend>��ݲ���</legend><a href='Admin_ActionList.asp?action=setgrade&Name="&Rs(0)&"'>�����ر�ȼ���</a><br><a href='admin_action.asp?action=BoardAdmin&Name="&Rs(0)&"'>��������</a><br><a href=#this onclick=""checkclick('ɾ�����û������������ӣ�\n\nɾ���󽫲��ָܻ�����ȷ��Ҫɾ����','Admin_Confirm.asp?Action=UpdateUserList&point=2&ID="&Rs(35)&"')"">��ȫɾ��</a><br><a href='Admin_Action.asp?action=AddLockIp&IP="&Rs(15)&"&Readme=��("&Rs(0)&")��ע��IP'>����ע��IP</a><br><a href='Admin_Action.asp?action=AddLockIp&IP="&Rs(17)&"&Readme=��("&Rs(0)&")������½IP'> ��������½IP</a></fieldset></div><fieldset><legend>�û���Ϣ</legend>���ڵȼ��飺"&BBS.GetGradeName(Rs(31),Rs(32))&"<br>ע���Աʱ�䣺"&Rs(14)&" <br>����½ʱ�䣺"&Rs(16)&" <br>ע���ԱʱIP��¼��"&Rs(15)&"<br>����½ʱIP��¼��"&Rs(17)&"</fieldset></div>"
	Response.Write"<div style='text-align:left; padding:3px' class='divth'><b>�û�ע����Ϣ</b></div>"
	DIVTR"�û����ƣ�","","<input name='ID' type='hidden' value='"&Rs(35)&"'><input type='text' class='text' name='Name' size='20' value='"&Rs(0)&"' />",25,1
	DIVTR"�û����룺","","<input type='text' class='text' name='Password' size='20' value='' /> �����벻Ҫ��",25,2
	DIVTR"�������⣺","","<input type='text' class='text' name='Clue' size='20' value='"&Rs(2)&"' />",25,1
	DIVTR"����𰸣�","","<input type='text' class='text' name='Answer' size='20' value='' /> �����벻Ҫ��",25,2
	DIVTR"�Ա�","",GetRadio("Sex","Ů",Rs(4),0)&GetRadio("Sex","��",Rs(4),1),25,1
	DIVTR"���䣺","","<input type='text' class='text' name='Mail' size='20' value='"&Rs(5)&"' />",25,2
	DIVTR"���գ�","","<input type='text' class='text' name='Birthday' size='20' value='"&Rs(6)&"' />",25,1
	DIVTR"��ҳ��","","<input type='text' class='text' name='Home' size='40' value='"&Rs(7)&"' />",25,2
	DIVTR"QQ���룺","","<input type='text' class='text' name='QQ' size='20' value='"&Rs(8)&"' />",25,1
	DIVTR"����QQ������Ϊͷ��","",GetRadio("isQQpic","��",Rs(9),0)&GetRadio("isQQpic","��",Rs(9),1)&"(QQ���������д)",25,2
	DIVTR"ͷ��","","<input type='text' class='text' name='Pic' size='40' value='"&Rs(10)&"' />",25,1
	DIVTR"ͷ��ߴ磺","","��<input type='text' class='text' name='PicW' size='5' value='"&Rs(11)&"' /> �ߣ�<input type='text' class='text' name='PicH' size='5' value='"&Rs(12)&"' />",25,2
	DIVTR"ǩ����","","<TEXTAREA name='sign' rows='4' style='width: 98%;'>"&Rs(13)&"</textarea>",60,1
	Response.Write"<div style='text-align:left; padding:3px' class='divth'><b>�û���̳��Ϣ</b></div>"
	DIVTR"��������","","<input type='text' class='text' name='EssayNum' size='6' value='"&Rs(18)&"' />",25,1
	DIVTR"����������","","<input type='text' class='text' name='GoodNum' size='6' value='"&Rs(19)&"' />",25,2
	DIVTR"���֣�","","<input type='text' class='text' name='Mark' size='6' value='"&Rs(20)&"' />",25,1
	DIVTR"��Ǯ��","","<input type='text' class='text' name='Coin' size='6' value='"&Rs(21)&"' />",25,2
	DIVTR"��","","<input type='text' class='text' name='BankSave' size='6' value='"&Rs(22)&"' />",25,1
	DIVTR"��Ϸ�ң�","","<input type='text' class='text' name='GameCoin' size='6' value='"&Rs(23)&"' />",25,2
	DIVTR"��½������","","<input type='text' class='text' name='LoginNum' size='6' value='"&Rs(26)&"' />",25,1
	DIVTR"ͷ�Σ�","","<input type='text' class='text' name='Honor' size='30' value='"&Rs(24)&"' />",25,2
	DIVTR"���ɣ�","","<input type='text' class='text' name='Faction' size='30' value='"&Rs(25)&"' />",25,1
	Response.Write"<div style='text-align:left; padding:3px' class='divth'><b>�����û�ѡ��</b></div>"
	DIVTR"��ʱɾ����","",GetRadio("isDel","��",Rs(27),0)&GetRadio("isDel","��",Rs(27),1),25,1
	DIVTR"VIP��Ա��","",GetRadio("isVip","��",Rs(28),0)&GetRadio("isVip","��",Rs(28),1),25,2
	DIVTR"�������ӣ�","",GetRadio("isShow","��",Rs(29),0)&GetRadio("isShow","��",Rs(29),1),25,1
	DIVTR"����ǩ����","",GetRadio("isSign","��",Rs(30),0)&GetRadio("isSign","��",Rs(30),1),25,2
	Response.Write"<div class='bottom'><input type='submit' class='button' value='�� ��'><input class='button' type='reset' value='�� ��'></div></div></form>"
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
		GoBack"","һЩ�������������д":Exit Sub
	End If
	Set Rs=BBS.Execute("select name,GradeID,GradeFlag From[User] where ID="&ID&"")
	If Rs.eof Then
		GoBack"","����û����������ڣ�":Exit Sub
	Else
		OldName=Rs(0)
		GradeID=Rs(1)
		GradeFlag=Rs(2)
	End If
	Rs.close
	If Password<>"" Then
		If len(password)<6 Then Goback"","���벻������6λ":Exit sub
		Password="[Password]='"&Md5(Password)&"',"
	End If
	If Answer<>"" Then Answer="Answer='"&Md5(Answer)&"',"
	If Isdate(Birthday) Then Birthday="Birthday='"&Birthday&"'," Else Birthday="Birthday=null,"
	
	If lcase(Name)<>Lcase(OldName) Then
		If Not BBS.Execute("select name From[User] where Name='"&Name&"' And ID<>"&ID&"").eof Then
			GoBack"","���û������Ѿ���ע����,���ܸ�����":Exit Sub
		End If
	End If
	Temp="update [User] Set "&Password&" Clue='"&Clue&"',"&Answer&"Sex="&Sex&",Mail='"&Mail&"',"&Birthday&" Home='"&Home&"',QQ='"&QQ&"',isQQpic="&isQQpic&",Pic='"&Pic&"',PicW="&PicW&",PicH="&PicH&",Sign='"&Sign&"',EssayNum="&EssayNum&",GoodNum="&GoodNum&",Mark="&Mark&",Coin="&Coin&",BankSave="&BankSave&",GameCoin="&GameCoin&",Honor='"&Honor&"',Faction='"&Faction&"',LoginNum="&LoginNum&",isDel="&isDel&",isVip="&isVip&",isShow="&isShow&",isSign="&isSign&" where ID="&ID
	BBS.Execute(Temp)
	OldName=Replace(OldName,"'","''")
	Temp="�����û���"&OldName&"��������!"
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
		Temp=Temp&" �û�����Ϊ��"&Name&"��"
	End If
'======>>>���µȼ�
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
	If GradeFlag=1 Then'���Ϊ������
		If BBS.Execute("Select ID From [grade] where ID="&GradeID).Eof Then GradeFlag=0
	End IF
	BBS.UpdateGrade ID,EssayNum,GradeFlag
'<<-----
	BBS.NetLog "������̨_"&Temp
	Suc"",Temp,"?Action=EditUser&ID="&ID
End Sub


%>