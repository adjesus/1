<!--#include file="Admin_Check.asp"-->
<!--#include file="Inc/page_Cls.asp"-->
<script language="JavaScript" type="text/javascript">
function showexplain(){
var ex=document.getElementById('explain').style
if (ex.display=='block')ex.display='none';
else ex.display='block'
}
</script>
<%
Dim PageInfo
Head()
Select case lcase(Request.querystring("Action"))
Case"log"
	CheckString "07"
	ShowLog
Case"placard"
	CheckString "03"
	Placard
Case"link"
	CheckString "05"
	Link
Case"userlist"
	CheckString "21"
	UserList
Case"recycle"
	CheckString "36"
	Recycle
Case"see"
	CheckString "36"
	See
Case"setgrade"
	CheckString "26"
	SetGrade
End select
Footer()

Function GetPageInfo(PTable,PFieldslist,PCondiction,POrderlist,PPrimaryKey,PSize,PCookiesName,Purl)
	Dim P
	Set P = New Cls_PageView
	P.strTableName =PTable
	P.strFieldsList =PFieldslist
	P.strCondiction =PCondiction
	P.strOrderList = POrderlist
	P.strPrimaryKey = PPrimaryKey
	P.intPageSize = PSize
	P.intPageNow = Request("page")
	P.strCookiesName = PCookiesName
	P.strPageUrl = PUrl
	P.InitClass
	GetPageInfo = P.arrRecordInfo
	PageInfo = P.strPageInfo
	Set P = nothing
End Function

Sub Link
	Dim Arr_Rs,I,Key,trColor,Sqlwhere,Flag,Chk,Pass
	Key=BBS.Fun.Getkey("Key")
	Flag=Request("Flag")
	Pass=Request("Pass")
	If Key<>"" Then
		Select Case Flag
			Case"1":Sqlwhere=BBS.Fun.SplitKey("BbsName",Key,"or")
			Case"2":Sqlwhere=BBS.Fun.SplitKey("Admin",Key,"or")
			Case"3":Sqlwhere=BBS.Fun.SplitKey("Readme",Key,"or")
			Case Else
			Sqlwhere=BBS.Fun.SplitKey("BbsName",Key,"or")&" or "&BBS.Fun.SplitKey("Admin",Key,"or")&" or "&BBS.Fun.SplitKey("Readme",Key,"or")
		End Select
	ElseIf Pass<>"" Then
		Sqlwhere="Pass="&Pass
	Else
		Sqlwhere=""
	End If
	Arr_rs=GetPageInfo("[Link]","ID,BbsName,Admin,Url,Orders,Ispic,pass,Readme,IsIndex",SqlWhere,"Orders","ID",20,"Link"&Key&Flag,"?Action=Link&Key="&Key&"&Flag="&Flag&"&Pass="&Pass)
	Response.Write"<div class='mian'><form method='post' style='margin:0px' action='?Action=Link'>"&_
	"<div class='top'>��̳����</div><div class='divtr2'><span style='float:right;padding:3px'><a href='Admin_Action.asp?Action=A_E_Link'>"&IconA&"�������</a> �鿴����<a href='?action=Link&pass=1'>�����</a>�� ��<a href='?action=Link&pass=0'>δ���</a>�� </span>&nbsp;������<input name='Key' class='text' value='"&Key&"'> <select name='Flag'><option value='1' selected>��̳����</option><option value='2'>��̳վ��</option><option value='3'>��̳���</option><option value='0'>�������</option></select><input type='submit'  class='button' value='����'></div></form>"
	Response.Write"<form method='post' style='margin:0px' action='Admin_Confirm.asp?Action=UpdateLink' onSubmit='ok.disabled=true;ok.value=""���ڸ���-���Եȡ�����""'>"
	If IsArray(Arr_Rs) Then
	Response.Write"<table class='Stable' border='0' cellpadding='3' cellspacing='0'>"&_
	"<tr><th>��̳����</th><th width=15%'>վ��</th><th width='10%'>����</th><th width='10%'>��ҳ</th><th width='10%'>ͼƬ</th><th width='10%'>���</th><th width='20%'>����</th></tr>"
		For i = 0 to UBound(Arr_Rs, 2)
			IF I mod 2 = 0  Then Response.Write"<tr >" Else Response.Write"<tr bgcolor='#DEF0FE'>"
			Response.Write "<td><a href='"&Arr_Rs(3,i)&"' target='_blank' title='"&Arr_Rs(7,i)&"'>"&BBS.Fun.ReplaceKey(Arr_Rs(1,i),Key)&"</a></td><td>"&BBS.Fun.ReplaceKey(Arr_Rs(2,I),Key)&"</td><td align='center'><input name='id' value='"&Arr_Rs(0,i)&"' type='hidden' /><input type='text' class='text' name='orders' value='"&Arr_Rs(4,i)&"' size='3' /></td>"
			If Arr_Rs(8,i)="1" Then Chk="checked" Else Chk=""
			Response.Write "<td align='center'><input name='isindex"&I+1&"' type='checkbox' value='1' "&Chk&"></td>"
			If Arr_Rs(5,i)="1" Then Chk="checked" Else Chk=""
			Response.Write "<td align='center'><input name='ispic"&I+1&"' type='checkbox' value='1' "&Chk&"></td>"
			If Arr_Rs(6,i)="1" Then Chk="checked" Else Chk=""
			Response.Write "<td align='center'><input name='pass"&I+1&"' type='checkbox' value='1' "&Chk&"></td><td><a href='Admin_Action.asp?Action=A_E_Link&ID="&Arr_Rs(0,i)&"'>"&IconE&"�༭</a> <a href=#this onclick=checkclick('ɾ���󽫲��ָܻ�����ȷ��Ҫɾ����','Admin_Confirm.asp?Action=DelLink&ID="&Arr_Rs(0,i)&"')>"&IconD&"ɾ��</td></tr>"
		Next
	Response.Write"</table><div class='bottom'><input type='Submit' name='ok' class='button' value='�������±�ҳ' /><input type='reset' class='button' value='�� ��' /></td></div></form>"&_
	"<div class='divth'>"&pageInfo&"</div>"
	Else
	Response.Write"<div class='divtr1'>û���ҵ�<span style='color:#F00'>"&Key&"</span>�ļ�¼��</div>"
	End If
	Response.Write"</div>"
End Sub



Sub ShowLog
	Dim Arr_Rs,I,Key,trColor,Sqlwhere,Flag
	Key=BBS.Fun.Getkey("Key")
	Flag=Request("Flag")
	If Key<>"" Then
		Select Case Flag
			Case"1":Sqlwhere=BBS.Fun.SplitKey("UserName",Key,"or")
			Case"2":Sqlwhere=BBS.Fun.SplitKey("Remark",Key,"or")
			Case"3":Sqlwhere=BBS.Fun.SplitKey("GetUrl",Key,"or")
			Case Else
			Sqlwhere=BBS.Fun.SplitKey("UserName",Key,"or")&" or "&BBS.Fun.SplitKey("Remark",Key,"or")&" or "&BBS.Fun.SplitKey("GetUrl",Key,"or")
		End Select
	Else
		Sqlwhere=""
	End If
	Arr_rs=GetPageInfo("[Log]","ID,Username,UserIP,Remark,logtime,GetUrl",SqlWhere,"ID desc","ID",20,"Log"&Key&Flag,"?Action=Log&Key="&Key&"&Flag="&Flag)
	Response.Write"<div class='mian'><form method='post' style='margin:0px' action='?Action=log'>"&_
	"<div class='top'>��̳��־ϵͳ</div><div class='divtr2'>������־ �ؼ��֣�<input name='Key' class='text' value='"&Key&"'> <select name='Flag'><option value='0'>ȫ��</option><option value='1'>������</option><option value='2' selected>�¼�����</option><option value='3'>��ַ����</option></select><input type='submit' class='button' value='����'></div></form>"
	Response.Write"<form method='post' style='margin:0px' action='Admin_Confirm.asp?Action=DelLog'>"
	If IsArray(Arr_Rs) Then
	Response.Write"<table class='Stable' border='0' cellpadding='3' cellspacing='0'>"&_
	"<tr><th width='10%'>������</th><th>�¼�����</th><th>��ַ����</th><th width='80'>ʱ��</td><th width='80'>IP</th><th width='28'>ѡ��</th></tr>"
		For i = 0 to UBound(Arr_Rs, 2)
			IF I mod 2 = 0  Then Response.Write"<tr >" Else Response.Write"<tr bgcolor='#DEF0FE'>"
			Response.Write "<td align='center'>"&BBS.Fun.ReplaceKey(Arr_Rs(1,i),Key)&"</td><td>"&BBS.Fun.ReplaceKey(Arr_Rs(3,I),Key)&"</td><td>"&BBS.Fun.ReplaceKey(Arr_Rs(5,I),Key)&"</td><td align='center'>"&Arr_Rs(4,I)&"</td><td align='center'>"&Arr_Rs(2,I)&"</td><td><input name='ID' type='checkbox' value='"&Arr_Rs(0,I)&"'></td></tr>"
		Next
	Response.Write"</table><div class='bottom'><input type='checkbox'  name='chkall' value='on' onClick='CheckAll(this.form)'>ȫѡ<input type='submit' class='button' value='ɾ����ѡ' name='Del'><input type='Submit' class='button' name='Del' value='�����־'></td></div></form>"&_
	"<div class='divth'>"&pageInfo&"</div>"
	Else
	Response.Write"<div class='divtr1'>û���ҵ�<span style='color:#F00'>"&Key&"</span>�ļ�¼��</div>"
	End If
	Response.Write"</div>"
End Sub


Sub Placard()
	Dim P,Page,arr_Rs,i,Temp,Content
	Arr_Rs=GetPageInfo("[Placard]","ID,Caption,BoardID,Name,AddTime,hits","","BoardID,ID desc","ID",20,"Placard_List","?Action=Placard")
	Response.Write"<div class='mian'><div class='top'><span style='float:right;padding:3px'><a href='admin_SetHtmlEdit.asp?Action=SayPlacard'>"&IconA&"<font color='#FFFFFF'>������</font></a></span>��̳����</div>"&_
	"<table border='0' class='Stable' cellpadding='3' cellspacing='0'><tr><th width='35%'>�������</th><th width='15%'>���ڰ��</th><th width='10%'>������</th><th width='10%'>ʱ��</th><th width='20%'>����</th></tr>"
	If IsArray(Arr_Rs) Then
		For i = 0 to UBound(Arr_Rs, 2)
		IF I mod 2 = 0  Then Response.Write"<tr >" Else Response.Write"<tr bgcolor='#DEF0FE'>"
			Response.Write"<td><a href='#this' onclick=""openwin('preview.asp?Action=placard&ID="&Arr_Rs(0,i)&"',500,400,'yes')"" >"&Arr_Rs(1,i)&"</a></td><td align='center'>"&GetBoardName(Arr_Rs(2,i))&"</td><td>"&Arr_Rs(3,i)&"</td><td>"&Arr_Rs(4,i)&"</td><td><a href='admin_SetHtmlEdit.asp?Action=SayPlacard&ID="&Arr_Rs(0,i)&"'>"&IconE&"�޸�</a> <a href=#this onclick=""checkclick('ɾ���������棡��\n\n��ȷ��Ҫɾ����','Admin_Confirm.asp?Action=delPlacard&ID="&Arr_Rs(0,i)&"')"">"&IconD&"ɾ��</a></td></tr>"
		Next	
	Response.Write"</table><div class='bottom'>"&PageInfo&"</div></div>"
	End If
End Sub

Function GetBoardName(Ast)
	Dim i
	If Not IsArray(BBS.Board_Rs) Then BBS.GetBoardCache()
	If IsArray(BBS.Board_Rs) Then
		For i=0 To Ubound(BBS.Board_Rs,2)
		IF BBS.Board_Rs(1,i)=Ast Then
			GetBoardName=BBS.Board_Rs(3,i)
			Exit For
		End IF
		Next
	End If
	If GetBoardName="" Then GetBoardName="��ҳ"
End Function

Function GradeList(Flag)
	Dim ARs,i
	If BBS.Cache.valid("GradeInfo") then
		ARs=BBS.Cache.Value("GradeInfo")
	Else
		ARs=BBS.SetGradeInfoCache()
	End if
	For i=0 To Ubound(ARs,2)
	If Flag=1 Then
		If ARs(1,i)="1" Then
			GradeList=GradeList&"<option value='"&ARs(0,i)&"'>"&ARs(2,i)&"</option>"
		End If
	Else
		GradeList=GradeList&"<option value='?action=Userlist&Flag=8&GradeID="&ARs(0,i)&"&GradeName="&ARs(2,i)&"'>"&ARs(2,i)&"</option>"
	End If
	Next
End Function


Sub SetGrade()
	Dim Name
	Name=Request("Name")
	Response.Write"<form method='post' style='margin:0px' action='Admin_Confirm.asp?Action=UpdateUserList'>"
	Response.Write"<div class='mian'><div class='top'>�����û��ر�ȼ���</div>"
	Response.Write"<div class='divtr1' style='padding:3px;'>�û�����<input name='Name' type='text' class='text' size='12' value='"&Name&"' />"
	Response.Write"<input name='point' type='radio' value='8' />����Ϊ�ر�ȼ��飺<select name='GradeID'>"&GradeList(1)&"</select> <input name='point' type='radio' value='9' /> ������ͨ�ȼ���(����������)</div>"
	Response.Write"<div class='bottom'><input type='submit' class='button' value='ȷ ��' /><input type='reset' class='button' value='�� ��' /></div></div></form>"
End Sub

Sub UserList
	Dim Arr_Rs,I,Key,trColor,Sqlwhere,Flag,Sex,Css,S
	Dim SqlSelect,SqlOrder,Title,TxtLink,GradeName,GradeID,Temp
	Flag=Request("Flag")
	GradeName=Replace(Replace(Request("GradeName"),"|",""),",","")
	GradeID=Request("GradeID")
	If Flag="" Then Flag="5"
	If Flag<>"8" Then GradeName="":GradeID=""
	SqlSelect=Split("0|ע�������û�|Isdel=2|,"&_
	"1|VIP�û�|IsVIP=1|Id Desc,"&_
	"2|��ɾ�����û�|IsDel=1|Id Desc,"&_
	"3|���������ӵ��û�|IsShow=1|Id Desc,"&_
	"4|������ǩ�����û�|IsSign=1|Id Desc,"&_
	"5|�����û�||Id Desc,"&_
	"6|���������û�||EssayNum desc,"&_
	"7|û�з������û�|EssayNum=0|Regtime,"&_
	"8|"&GradeName&"|GradeID="&GradeID&"|ID Desc",",")
	Txtlink="���ܲ�����"
	For i=0 To uBound(SqlSelect)
		If i="5" then TxtLink=Txtlink&"<br>���ٲ鿴��"
		If i="8" Then Txtlink=Txtlink&" <select onchange=if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;} style='font-size: 9pt'><option selected>���ȼ���鿴</option>"&GradeList(0)&"</select>"
		Temp=Split(SqlSelect(i),"|")
		If Flag=Temp(0) Then
			Txtlink=Txtlink&" <a href='?Action=UserList&Flag="&Temp(0)&"&GradeName="&GradeName&"&GradeID="&GradeID&"' style='color:#F00'>"&Temp(1)&"</a> |"
			Title=Temp(1)
			Sqlwhere=Temp(2)
			SqlOrder=Temp(3)
		Else
			Txtlink=Txtlink&" <a href='?Action=UserList&Flag="&Temp(0)&"&GradeName="&GradeName&"&GradeID="&GradeID&"'>"&Temp(1)&"</a> |"
		End If
	Next
	Response.Write"<div class='mian'><div class='top'>�û�����</div><div class='divtr2' style='padding:4px'>"&Txtlink&"</div></div>"
	Key=BBS.Fun.Getkey("Key")
	If Key<>"" Then
		Sqlwhere=BBS.Fun.SplitKey("Name",Key,"or")
	End If
	Arr_rs=GetPageInfo("[User]","ID,Name,Sex,EssayNum,LastIp,Lasttime,Mail,GoodNum,Mark,Home,QQ,GradeID,Coin,BankSave,Sign,Pic,PicW,PicH,Birthday,BankTime,Regtime,NewSmsNum,SmsSize,isQQpic,isShow,isDel,isVip,isSign,RegIp,LoginNum,Honor,Faction,GameCoin",Sqlwhere,SqlOrder,"ID",25,"UserList"&Key&Flag,"?Action=UserList&Key="&Key&"&GradeName="&GradeName&"&GradeID="&GradeID&"&Flag="&Flag)
	Response.Write "<div class='mian'><form method='post' style='margin:0px' action='?Action=UserList'>"&_
	"<div class='top'>"&Title&"</div><div class='divth'>���������û���<input name='Key' class='text' value='"&Replace(Key,"[[]","[")&"'><input type='submit' class='button' value='����' /></div></form>"
	Response.Write"<form method='post' style='margin:0px' action='Admin_Confirm.asp?Action=UpdateUserList'>"
	If IsArray(Arr_Rs) Then
	Response.Write"<div class='divtr2'>"&Replace(pageInfo,"����¼","λ�û�")&"</div><table class='Stable' border='0' cellpadding='3' cellspacing='0'>"&_
	"<tr><th width='27'>ѡ��</th><th>�û�����(����༭)</th><th width='30'>�Ա�</th><th>����</th><th width='20%'>���IP(�������)</th><th width='120'>����½</td><th>Email(�������)</th></tr>"
		For i = 0 to UBound(Arr_Rs, 2)
	If Arr_Rs(2,i)=1 Then Sex="��" Else Sex="Ů"
			IF I mod 2 = 0  Then Response.Write"<tr >" Else Response.Write"<tr bgcolor='#DEF0FE'>"
			Response.Write "<td><input name='ID' type='checkbox' id='ID' value='"&Arr_Rs(0,I)&"'></td><td><a href='Admin_user.asp?Action=EditUser&ID="&Arr_Rs(0,i)&"'>"&IconE&BBS.Fun.ReplaceKey(Arr_Rs(1,i),Key)&"</a></td><td align='center'>"&Sex&"</td><td align='center'>"&Arr_Rs(3,I)&"</td><td><a href='Admin_Action.asp?action=AddLockIp&IP="&Arr_Rs(4,I)&"&Readme=�û�����"&Arr_Rs(1,I)&"' title='�����û�IP'><img src='Images/icon/lock.gif' border='0' alt='����IP' /> "&Arr_Rs(4,I)&"</a></td><td align='center'>"&Arr_Rs(5,I)&"</td><td><a href='mailto:"&arr_rs(6,i)&"'>"&arr_rs(6,i)&"</a></td></tr>"
	Next
	Response.Write"</table><div class='divtr2'><div class='divtd1' style='width:50px'>"&_
	"<input type='checkbox' name='chkall' value='on' onClick='CheckAll(this.form)' />ȫѡ</div>"&_
	"<div class='divtd2'>&nbsp;&nbsp;������"
	Temp=""
	If Flag="0" Then
		Temp="<input name='point' type='radio' value='10' checked />ͨ����� "
	Else
		If Flag="2" Then 
			Temp="<input name='point' type='radio' value='12' checked />�ָ��û� "
		Else
			S=S&"<input name='point' type='radio' value='1' />��ʱɾ��(���Իָ�) "
		End If
		If Flag="3" Then
			Temp="<input name='point' type='radio' value='13' checked />�ָ������� "
		Else
			S=S&"<input name='point' type='radio' value='4' />���������� "
		End If
		If Flag="4" Then
			Temp="<input name='point' type='radio' value='14' checked />�ָ���ǩ�� "
		Else
			S=S&"<input name='point' type='radio' value='5'/>������ǩ�� "
		End If
		If Flag="1" Then
			Temp="<input name='point' type='radio' value='11' checked />ȡ��VIP "
		Else
			S=S&"<input name='point' type='radio' value='6' />����ΪVIP "
		End If
	End If
	Response.Write Temp&" <input name='point' type='radio' value='8' />�����ر�ȼ��飺<select name='GradeID'>"&GradeList(1)&"</select> "
	If Flag="8" Then Response.Write"<input name='point' type='radio' value='9' /> ȡ���ر�ȼ���"
	Response.Write "<br />"&S&"<input name='point' type='radio' value='2' />��ȫɾ��(��������) <input name='point' type='radio' value='3' />ɾ��������  <input name='point' type='radio' value='7' />�޸�</div>"
	Response.Write"<div style='clear:both'></div></div><div class='bottom'><input type='submit' class='button' value='����ִ�в���' /><input class='button' type='button' value='˵��' onclick='showexplain()' /></div></div></form>"
	Else
	Response.Write"<div class='divtr1'>û���ҵ�<span style='color:#F00'>"&Key&"</span> ��¼��</div>"
	End If
	Response.Write"</div>"
Response.Write"<div class='mian' id='explain' style='display:none'><div class='top'>����˵��</div><div class='divtr2' style='padding:5px'><li>�����ر�ȼ��飺ѡ�еĻ�Ա�����󣬲����������ƣ���������һЩ��̳Ȩ�ޡ�</li><li>��ʱɾ������ʱɾ��ѡ����û���ֻ����ǣ�������ʱ�ָ���</li><li>��ȫɾ������ȫɾ��ѡ����û������������Ӻ����Եȣ�ɾ���󲻿ɻָ���</li><li>ɾ�����ӣ�ɾ��ѡ����û����������ӣ�</li><li>�������ӣ�����ѡ����û����������ӣ����ӽ�������ʾ��</li><li>����ǩ��������ѡ����û���ǩ�������Ӳ�����ʾǩ������ </li><li>����ΪVIP��ֱ�ӽ�ѡ����û�����ΪVIP��Ա�� </li><li>�޸��������޸�ѡ���û��ĵȼ��������������������ȡ�</li></div></div>"
End Sub
%>