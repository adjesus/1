<!--#include file="Inc.asp"-->
<!--#include file="Inc/md5.asp"-->
<%
Dim action
action=Lcase(Request.Querystring("action"))
If Len(action)>15 Then BBS.GotoErr(1)
BBS.CheckMake
if action<>"getpassword" and action<>"forgetpassword" Then
	BBS.Position=BBS.Position&" -> <a href=""userinfo.asp"">�û��������<a>"
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
	Response.Write BBS.ReadSkins("�û��������")
End Sub

Sub Myinfo()
	Dim Rs,S,temp,Temp1
	BBS.Head "","","�޸ĸ�������"
	MyManager()
	SET RS=BBS.Execute("Select Name,Sex,Birthday,Mail,Home,IsQQpic,QQ,Pic,Pich,Picw,IsSign,Sign,Honor,GradeID From[user]where ID="&BBS.MyID&" And Isdel=False")
	IF Rs.eof Then BBS.MakeCookiesEmpty():BBS.GotoErr(100)
	S="<form style='margin:0' method='POST' action='?action=savemyinfo'  name='form'>"
	S=S&"<div style='text-align:left;border-bottom:1px solid "&BBS.SkinsPIC(0)&";padding:3px;font-weight:bold;BACKGROUND: "&BBS.SkinsPIC(2)&";'><img src='Images/icon/inn.gif' align='absmiddle' atl='��������' /> ��������</div>"
	S=S&BBS.Row("<b>�û�����</b>��<br>��Ϊ��̳ID�ʺţ��Լ������޸�","<br /><b>"&Rs("Name")&"</b>","65%","45px")
	If Rs("Sex")=1 Then Temp=" checked" Else Temp=""
	If Rs("Sex")=0 Then Temp1=" checked" Else Temp1=""
	S=S&BBS.Row("<b>�����Ա�</b>","<input name='sex' type='radio' value='1' "&Temp&" class=checkbox /><img src='Images/icon/male.gif' align='absmiddle' atl='˧��' /> ˧��&nbsp;&nbsp;<input type='radio' name='sex' value='0' "&Temp1&" class=checkbox /><img src='Images/icon/female.gif' align='absmiddle' atl='��Ů' /> ��Ů","65%","")
	S=S&BBS.Row("<b>Email��ַ</b>��<br>��������Ч���ʼ���ַ","<input type='text' class='text' name='mail' style='margin-top:8px' size='30' maxlength='30' value='"&Rs("Mail")&"' />","65%","44px")
	S=S&"<div style='text-align:left;border-bottom:1px solid "&BBS.SkinsPIC(0)&";padding:3px;font-weight:bold;BACKGROUND: "&BBS.SkinsPIC(2)&";'><img src='Images/icon/inn.gif' align='absmiddle' atl='ѡ������' /> ѡ������</div>"
	S=S&"<div id='Rowhow' display='none'>"
	S=S&BBS.Row("<b>���գ�</b>","<input  type='text' class='text' name='birthday' id='birthday' size='10' maxlength='10' readonly='true' onfocus=""show_cele_date('birthday','','',this)"" value='"&Rs("Birthday")&"' />","65%","20px")
	S=S&BBS.Row("<b>��ҳ��</b><br />��д��ĸ�����ҳ���ô�Ҽ�ʶ��ʶ��","<input type='text' class='text' name='home' size='30' maxlength='200'  style='margin-top:8px' value='"&Rs("Home")&"' />","65%","44px")
	If Rs("IsQQPic") Then
	Temp=" checked"
	Else
	Temp=""
	End If
	S=S&BBS.Row("<b>QQ���룺</b><br />��д����QQ��ַ�����������˵���ϵ","<input type='text' class='text'  style='margin-top:8px' value='"&Rs("QQ")&"'name='QQ'  maxlength='15'> <input type='checkbox' onclick='QQpic()' name='isqqpic' id='isqqpic' value='1' "&Temp&" class=checkbox />����QQ������Ϊͷ��","65%","44px")
	S=S&"<div id='showpic'>"
	S=S&BBS.Row("<b>ѡ����̳ͷ��</b><br />ʹ����̳�Դ���ͼ��",HeadPicOpt() &"<img src='"&Rs("Pic")&"' id='pic' name='pic' /> <input onclick=""openwin('preview.asp?action=HeadPic',680,400,'yes')"" type='button' class='button' value='ȫ��ͷ��' />","65%","50")
	If SESSION(CacheName& "MyGradeinfo")(14)="1" then
	Temp="<input type=button value='�ϴ�ͷ��ͼƬ' class='button' onclick=""javascript:up.style.display='block';upf.location.href='UploadFile.asp?Flag=1';this.style.display='none'""><div id='up' style='display:none'><iframe id='upf' scrolling='no' frameborder='0' height='22' width='100%'></iframe></div><br>"
	Else
	Temp=""
	End if
	S=S&BBS.Row("<b>�Զ���ͷ��</b><br />���ͼ��λ����������ͼƬ�����Զ����Ϊ��",Temp&"<input id='picurl' name='picurl' size='40' maxlength='100'  value='"&Rs("Pic")&"' /> ����Url��ַ<br />ͼ���ȣ�<input type='text' class='text' name='picw' id='picw'  size='6' value='"&Rs("PicW")&"' /> �߶ȣ�<input type=text name='pich' id='pich' size='6'  value='"&Rs("PicH")&"' />(����޶�:120)","65%","")
	S=S&"</div>"
	If SESSION(CacheName& "MyGradeinfo")(8)="0" then
		 Temp="����û�ﵽ�����Զ�ͷ�γƺŵĵȼ�Ȩ��<input type='hidden' value='"&Rs("Honor")&"' name='Honor'>"
	Else
		 Temp="<input style='margin-top:8px' type='text' value='"&Rs("Honor")&"' name='Honor'>"
	End If
	S=S&BBS.Row("<b>�Զ���ͷ�γƺţ�</b><br />���8������",Temp,"65%","40px")
	S=S&BBS.Row("<b>����ǩ����</b><br />���ֽ�����������������µĽ�β��<br />�������ĸ���(���255���ַ�)","<TEXTAREA name='sign' rows='4' style='width: 98%;'>"&Rs("Sign")&"</textarea>","65%","")
	S=S&"</div><div style=""padding:5px;BACKGROUND: "&BBS.SkinsPIC(2)&";""  align='center'><input type='submit' value='ȷ���޸ģ�' class='button' />&nbsp;&nbsp; <input type='reset' class='button' value='ȡ����д��'  /></div></form>"
	BBS.ShowTable"�޸��ҵ�����",S
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
	.Head"","","�޸ĸ�������"
	MyManager()
	Dim Temp,Content,Rs,Sql,Name,Mail,PicUrl,HeadPic,PicW,PicH,Home,Sign,QQ,IsQQpic,Sex,Birthday,Honor
	Name=.Fun.GetStr("name")
	Mail=.Fun.GetStr("mail")
	Honor=.Fun.GetStr("Honor")
	If Mail="" Then .GoToErr(42)
	Mail=server.HTMLEnCode(Mail)
	If Not .Fun.IsValidEmail(Mail) Then .GoToErr(42)
	'ֻ��һ������
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
	Sign=Replace(Left(.Fun.Replacehtml(.Fun.GetStr("Sign")),255),"{��������}","")
	PicH=.Fun.GetStr("pich")
	PicW=.Fun.Getstr("picw")
	If .info(57)="1" And (Instr(PicUrl,"://")>0  Or Instr(Lcase(Picurl),"www")>0 Or Instr(Lcase(PicUrl),"..")>0) Then  .GotoErr(45)'��ֹ�ⲿͼƬ
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
	Content="<div style='margin:15px;line-height: 150%'><li>�����޸ĳɹ���<li><a href=userinfo.asp>�����ҵ��û��������</a><li><a href=index.asp>������ҳ</a></div>"
	.ShowTable"�޸ĳɹ�",Content
	Session(CacheName & "Myinfo") = Empty
	End with
End Sub

Sub MyPassword
	Dim S
	BBS.Head "","","�޸�����"
	MyManager()
	S="<form style='margin:0' method='POST' action='?action=savemypassword' name='form'>"
	S=S&BBS.Row("<b>������ȷ�ϣ�</b><br />��������������ȷ��","<input type='password' name='Password' size='30' maxlength='20' class='text' />","65%","44px")
	S=S&BBS.Row("<b>�µ�����(���14λ)��</b><br />��ʹ�ó���'���͡�|���Լ�����������ַ�","<input type='password' name='NewPassword' size='30' maxlength='20' class='text' />","65%","44px")
	S=S&BBS.Row("<b>�ظ����룺</b><br />������һ��ȷ��","<input type='password' name='RePassword' size='30' maxlength='20' class='text' />","65%","44px")
	S=S&BBS.Row("<b>��������</b>��<br />�����������ʾ����","<input type='text' class='text' name='clue' size=30  maxlength='60' /> �粻���벻Ҫ��д","65%","40px")
	S=S&BBS.Row("<b>�����</b>��<br />�����������ʾ����𰸣�����ȡ����̳����","<input type='text' class='text' name='answer' size=30  maxlength='60' /> ͬ��","65%","40px")
	S=S&"<div style=""padding:5px;BACKGROUND: "&BBS.SkinsPIC(2)&";""  align='center'><input type='submit' value='ȷ���޸ģ�' class='button' />&nbsp;&nbsp; <input type='reset' class='button' value='ȡ�����ã�'  /></div></form>"
	BBS.ShowTable "�޸�����",S
End Sub

Sub savemypassword
	Dim Password,NewPassword,RePassword,Caption,Content,Clue,Answer
	BBS.CheckMake'��ֹ�ⲿ�ύ
	BBS.Head"","","�޸�����"
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
	Content="<div style='margin:15px;line-height: 150%'><li>�����޸ĳɹ�!</li><li><a href=userinfo.asp>�����û��������</a></li><li><a href=index.asp>������ҳ</a></li></div>"
	BBS.ShowTable "�޸ĳɹ�",Content
End Sub

Sub ForgetPassword
	Dim UserName,rs,S
	BBS.Head"","","�һ�����"
	UserName=BBS.Fun.GetStr("UserName")
	If UserName="" Then
	S="<form style='margin:0' method='POST'>"
	S=S&BBS.Row("<b>�����û�����</b>","<input type='text' name='UserName' size='30' maxlength='20' class='text' /> ������ע����û����ƽ���ȷ��","90%","22px")
	S=S&"<div style=""padding:5px;BACKGROUND: "&BBS.SkinsPIC(2)&";""  align='center'><input type='submit' value='ȷ����' class='button' /></div></form>"
	Else
	BBS.CheckMake'��ֹ�ⲿ�ύ
	If BBS.SafeBuckler(UserName,BBS.MyIP,0) Then BBS.Alert"BBS��ȫ����������\n\n�Բ����㳢���һ�������󳬹�3�Σ����첻�����һ������ˡ�\n�����Ϣ�ѱ�ϵͳ��¼��","Index.asp"
	set rs=BBS.Execute("select clue from [User] where name='"&UserName&"'")
	S="<form style='margin:0' method='POST' action='?action=getpassword'>"
	S=S&BBS.Row("<b>�û����ƣ�</b>","<input name='UserName' type='hidden' value='"&UserName&"'><b>"& UserName &"</b>","90%","22px")
	S=S&BBS.Row("<b>�������⣺</b>", ""&Rs("clue")&"","90%","22px")
	S=S&BBS.Row("<b>����𰸣�</b>","<input type='text' name='Answer' size='30' maxlength='20' class='text' /> ����������ע��ʱ��д�������","90%","22px")
	S=S&"<div style=""padding:5px;BACKGROUND: "&BBS.SkinsPIC(2)&";""  align='center'><input type='submit' value='ȷ����' class='button' />&nbsp;&nbsp; <input type='reset' class='button' value='���ã�'  /></div></form>"
	End If
	BBS.ShowTable"�һ�����",S
End Sub

Sub GetPassword
	Dim UserName,Clue,Answer,NewPassword,Content
	BBS.Head"","","�һ�����"
	UserName=BBS.Fun.GetStr("UserName")
	If BBS.SafeBuckler(UserName,BBS.MyIP,0) Then BBS.Alert"BBS��ȫ����������\n\n�Բ����㳢���һ�������󳬹�3�Σ����첻�����һ������ˡ�\n�����Ϣ�ѱ�ϵͳ��¼��","Index.asp"
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
		Content="<div style='margin:15px;line-height: 150%'><li>���ɹ���ͨ�����뱣���ļ��飡</li><li>�û����ƣ�<font color=red>"&UserName&"</font> &nbsp; ��������룺<font color=red>"&NewPassword&"</font></li><li>�ȼ�ס�����룬�������ϵ�½��̳�������޸����룡</li></div>"
		BBS.ShowTable "�ɹ�ͨ����֤",Content
	End If	
End Sub
%>
