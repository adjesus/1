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
	"<div class='top'>论坛联盟</div><div class='divtr2'><span style='float:right;padding:3px'><a href='Admin_Action.asp?Action=A_E_Link'>"&IconA&"添加连盟</a> 查看：【<a href='?action=Link&pass=1'>已审核</a>】 【<a href='?action=Link&pass=0'>未审核</a>】 </span>&nbsp;搜索：<input name='Key' class='text' value='"&Key&"'> <select name='Flag'><option value='1' selected>论坛名称</option><option value='2'>论坛站长</option><option value='3'>论坛简介</option><option value='0'>三项均搜</option></select><input type='submit'  class='button' value='搜索'></div></form>"
	Response.Write"<form method='post' style='margin:0px' action='Admin_Confirm.asp?Action=UpdateLink' onSubmit='ok.disabled=true;ok.value=""正在更新-请稍等。。。""'>"
	If IsArray(Arr_Rs) Then
	Response.Write"<table class='Stable' border='0' cellpadding='3' cellspacing='0'>"&_
	"<tr><th>论坛名称</th><th width=15%'>站长</th><th width='10%'>排序</th><th width='10%'>首页</th><th width='10%'>图片</th><th width='10%'>审核</th><th width='20%'>操作</th></tr>"
		For i = 0 to UBound(Arr_Rs, 2)
			IF I mod 2 = 0  Then Response.Write"<tr >" Else Response.Write"<tr bgcolor='#DEF0FE'>"
			Response.Write "<td><a href='"&Arr_Rs(3,i)&"' target='_blank' title='"&Arr_Rs(7,i)&"'>"&BBS.Fun.ReplaceKey(Arr_Rs(1,i),Key)&"</a></td><td>"&BBS.Fun.ReplaceKey(Arr_Rs(2,I),Key)&"</td><td align='center'><input name='id' value='"&Arr_Rs(0,i)&"' type='hidden' /><input type='text' class='text' name='orders' value='"&Arr_Rs(4,i)&"' size='3' /></td>"
			If Arr_Rs(8,i)="1" Then Chk="checked" Else Chk=""
			Response.Write "<td align='center'><input name='isindex"&I+1&"' type='checkbox' value='1' "&Chk&"></td>"
			If Arr_Rs(5,i)="1" Then Chk="checked" Else Chk=""
			Response.Write "<td align='center'><input name='ispic"&I+1&"' type='checkbox' value='1' "&Chk&"></td>"
			If Arr_Rs(6,i)="1" Then Chk="checked" Else Chk=""
			Response.Write "<td align='center'><input name='pass"&I+1&"' type='checkbox' value='1' "&Chk&"></td><td><a href='Admin_Action.asp?Action=A_E_Link&ID="&Arr_Rs(0,i)&"'>"&IconE&"编辑</a> <a href=#this onclick=checkclick('删除后将不能恢复！您确定要删除吗？','Admin_Confirm.asp?Action=DelLink&ID="&Arr_Rs(0,i)&"')>"&IconD&"删除</td></tr>"
		Next
	Response.Write"</table><div class='bottom'><input type='Submit' name='ok' class='button' value='批量更新本页' /><input type='reset' class='button' value='重 置' /></td></div></form>"&_
	"<div class='divth'>"&pageInfo&"</div>"
	Else
	Response.Write"<div class='divtr1'>没有找到<span style='color:#F00'>"&Key&"</span>的记录！</div>"
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
	"<div class='top'>论坛日志系统</div><div class='divtr2'>搜索日志 关键字：<input name='Key' class='text' value='"&Key&"'> <select name='Flag'><option value='0'>全部</option><option value='1'>操作人</option><option value='2' selected>事件内容</option><option value='3'>地址参数</option></select><input type='submit' class='button' value='搜索'></div></form>"
	Response.Write"<form method='post' style='margin:0px' action='Admin_Confirm.asp?Action=DelLog'>"
	If IsArray(Arr_Rs) Then
	Response.Write"<table class='Stable' border='0' cellpadding='3' cellspacing='0'>"&_
	"<tr><th width='10%'>操作人</th><th>事件内容</th><th>地址参数</th><th width='80'>时间</td><th width='80'>IP</th><th width='28'>选择</th></tr>"
		For i = 0 to UBound(Arr_Rs, 2)
			IF I mod 2 = 0  Then Response.Write"<tr >" Else Response.Write"<tr bgcolor='#DEF0FE'>"
			Response.Write "<td align='center'>"&BBS.Fun.ReplaceKey(Arr_Rs(1,i),Key)&"</td><td>"&BBS.Fun.ReplaceKey(Arr_Rs(3,I),Key)&"</td><td>"&BBS.Fun.ReplaceKey(Arr_Rs(5,I),Key)&"</td><td align='center'>"&Arr_Rs(4,I)&"</td><td align='center'>"&Arr_Rs(2,I)&"</td><td><input name='ID' type='checkbox' value='"&Arr_Rs(0,I)&"'></td></tr>"
		Next
	Response.Write"</table><div class='bottom'><input type='checkbox'  name='chkall' value='on' onClick='CheckAll(this.form)'>全选<input type='submit' class='button' value='删除所选' name='Del'><input type='Submit' class='button' name='Del' value='清空日志'></td></div></form>"&_
	"<div class='divth'>"&pageInfo&"</div>"
	Else
	Response.Write"<div class='divtr1'>没有找到<span style='color:#F00'>"&Key&"</span>的记录！</div>"
	End If
	Response.Write"</div>"
End Sub


Sub Placard()
	Dim P,Page,arr_Rs,i,Temp,Content
	Arr_Rs=GetPageInfo("[Placard]","ID,Caption,BoardID,Name,AddTime,hits","","BoardID,ID desc","ID",20,"Placard_List","?Action=Placard")
	Response.Write"<div class='mian'><div class='top'><span style='float:right;padding:3px'><a href='admin_SetHtmlEdit.asp?Action=SayPlacard'>"&IconA&"<font color='#FFFFFF'>发表公告</font></a></span>论坛公告</div>"&_
	"<table border='0' class='Stable' cellpadding='3' cellspacing='0'><tr><th width='35%'>公告标题</th><th width='15%'>所在版块</th><th width='10%'>发布者</th><th width='10%'>时间</th><th width='20%'>管理</th></tr>"
	If IsArray(Arr_Rs) Then
		For i = 0 to UBound(Arr_Rs, 2)
		IF I mod 2 = 0  Then Response.Write"<tr >" Else Response.Write"<tr bgcolor='#DEF0FE'>"
			Response.Write"<td><a href='#this' onclick=""openwin('preview.asp?Action=placard&ID="&Arr_Rs(0,i)&"',500,400,'yes')"" >"&Arr_Rs(1,i)&"</a></td><td align='center'>"&GetBoardName(Arr_Rs(2,i))&"</td><td>"&Arr_Rs(3,i)&"</td><td>"&Arr_Rs(4,i)&"</td><td><a href='admin_SetHtmlEdit.asp?Action=SayPlacard&ID="&Arr_Rs(0,i)&"'>"&IconE&"修改</a> <a href=#this onclick=""checkclick('删除这条公告！！\n\n您确定要删除吗？','Admin_Confirm.asp?Action=delPlacard&ID="&Arr_Rs(0,i)&"')"">"&IconD&"删除</a></td></tr>"
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
	If GetBoardName="" Then GetBoardName="首页"
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
	Response.Write"<div class='mian'><div class='top'>设置用户特别等级组</div>"
	Response.Write"<div class='divtr1' style='padding:3px;'>用户名：<input name='Name' type='text' class='text' size='12' value='"&Name&"' />"
	Response.Write"<input name='point' type='radio' value='8' />提升为特别等级组：<select name='GradeID'>"&GradeList(1)&"</select> <input name='point' type='radio' value='9' /> 降回普通等级组(按发帖计算)</div>"
	Response.Write"<div class='bottom'><input type='submit' class='button' value='确 定' /><input type='reset' class='button' value='重 置' /></div></div></form>"
End Sub

Sub UserList
	Dim Arr_Rs,I,Key,trColor,Sqlwhere,Flag,Sex,Css,S
	Dim SqlSelect,SqlOrder,Title,TxtLink,GradeName,GradeID,Temp
	Flag=Request("Flag")
	GradeName=Replace(Replace(Request("GradeName"),"|",""),",","")
	GradeID=Request("GradeID")
	If Flag="" Then Flag="5"
	If Flag<>"8" Then GradeName="":GradeID=""
	SqlSelect=Split("0|注册待审的用户|Isdel=2|,"&_
	"1|VIP用户|IsVIP=1|Id Desc,"&_
	"2|被删除的用户|IsDel=1|Id Desc,"&_
	"3|被屏蔽帖子的用户|IsShow=1|Id Desc,"&_
	"4|被屏蔽签名的用户|IsSign=1|Id Desc,"&_
	"5|所有用户||Id Desc,"&_
	"6|发帖最多的用户||EssayNum desc,"&_
	"7|没有发帖的用户|EssayNum=0|Regtime,"&_
	"8|"&GradeName&"|GradeID="&GradeID&"|ID Desc",",")
	Txtlink="功能操作："
	For i=0 To uBound(SqlSelect)
		If i="5" then TxtLink=Txtlink&"<br>快速查看："
		If i="8" Then Txtlink=Txtlink&" <select onchange=if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;} style='font-size: 9pt'><option selected>按等级组查看</option>"&GradeList(0)&"</select>"
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
	Response.Write"<div class='mian'><div class='top'>用户管理</div><div class='divtr2' style='padding:4px'>"&Txtlink&"</div></div>"
	Key=BBS.Fun.Getkey("Key")
	If Key<>"" Then
		Sqlwhere=BBS.Fun.SplitKey("Name",Key,"or")
	End If
	Arr_rs=GetPageInfo("[User]","ID,Name,Sex,EssayNum,LastIp,Lasttime,Mail,GoodNum,Mark,Home,QQ,GradeID,Coin,BankSave,Sign,Pic,PicW,PicH,Birthday,BankTime,Regtime,NewSmsNum,SmsSize,isQQpic,isShow,isDel,isVip,isSign,RegIp,LoginNum,Honor,Faction,GameCoin",Sqlwhere,SqlOrder,"ID",25,"UserList"&Key&Flag,"?Action=UserList&Key="&Key&"&GradeName="&GradeName&"&GradeID="&GradeID&"&Flag="&Flag)
	Response.Write "<div class='mian'><form method='post' style='margin:0px' action='?Action=UserList'>"&_
	"<div class='top'>"&Title&"</div><div class='divth'>快速搜索用户：<input name='Key' class='text' value='"&Replace(Key,"[[]","[")&"'><input type='submit' class='button' value='搜索' /></div></form>"
	Response.Write"<form method='post' style='margin:0px' action='Admin_Confirm.asp?Action=UpdateUserList'>"
	If IsArray(Arr_Rs) Then
	Response.Write"<div class='divtr2'>"&Replace(pageInfo,"条记录","位用户")&"</div><table class='Stable' border='0' cellpadding='3' cellspacing='0'>"&_
	"<tr><th width='27'>选择</th><th>用户名称(点击编辑)</th><th width='30'>性别</th><th>帖数</th><th width='20%'>最后IP(点击封锁)</th><th width='120'>最后登陆</td><th>Email(点击发送)</th></tr>"
		For i = 0 to UBound(Arr_Rs, 2)
	If Arr_Rs(2,i)=1 Then Sex="男" Else Sex="女"
			IF I mod 2 = 0  Then Response.Write"<tr >" Else Response.Write"<tr bgcolor='#DEF0FE'>"
			Response.Write "<td><input name='ID' type='checkbox' id='ID' value='"&Arr_Rs(0,I)&"'></td><td><a href='Admin_user.asp?Action=EditUser&ID="&Arr_Rs(0,i)&"'>"&IconE&BBS.Fun.ReplaceKey(Arr_Rs(1,i),Key)&"</a></td><td align='center'>"&Sex&"</td><td align='center'>"&Arr_Rs(3,I)&"</td><td><a href='Admin_Action.asp?action=AddLockIp&IP="&Arr_Rs(4,I)&"&Readme=用户名："&Arr_Rs(1,I)&"' title='封锁用户IP'><img src='Images/icon/lock.gif' border='0' alt='封锁IP' /> "&Arr_Rs(4,I)&"</a></td><td align='center'>"&Arr_Rs(5,I)&"</td><td><a href='mailto:"&arr_rs(6,i)&"'>"&arr_rs(6,i)&"</a></td></tr>"
	Next
	Response.Write"</table><div class='divtr2'><div class='divtd1' style='width:50px'>"&_
	"<input type='checkbox' name='chkall' value='on' onClick='CheckAll(this.form)' />全选</div>"&_
	"<div class='divtd2'>&nbsp;&nbsp;操作："
	Temp=""
	If Flag="0" Then
		Temp="<input name='point' type='radio' value='10' checked />通过审核 "
	Else
		If Flag="2" Then 
			Temp="<input name='point' type='radio' value='12' checked />恢复用户 "
		Else
			S=S&"<input name='point' type='radio' value='1' />暂时删除(可以恢复) "
		End If
		If Flag="3" Then
			Temp="<input name='point' type='radio' value='13' checked />恢复其帖子 "
		Else
			S=S&"<input name='point' type='radio' value='4' />屏蔽其帖子 "
		End If
		If Flag="4" Then
			Temp="<input name='point' type='radio' value='14' checked />恢复其签名 "
		Else
			S=S&"<input name='point' type='radio' value='5'/>屏蔽其签名 "
		End If
		If Flag="1" Then
			Temp="<input name='point' type='radio' value='11' checked />取消VIP "
		Else
			S=S&"<input name='point' type='radio' value='6' />提升为VIP "
		End If
	End If
	Response.Write Temp&" <input name='point' type='radio' value='8' />提升特别等级组：<select name='GradeID'>"&GradeList(1)&"</select> "
	If Flag="8" Then Response.Write"<input name='point' type='radio' value='9' /> 取消特别等级组"
	Response.Write "<br />"&S&"<input name='point' type='radio' value='2' />完全删除(包括帖子) <input name='point' type='radio' value='3' />删除其帖子  <input name='point' type='radio' value='7' />修复</div>"
	Response.Write"<div style='clear:both'></div></div><div class='bottom'><input type='submit' class='button' value='批量执行操作' /><input class='button' type='button' value='说明' onclick='showexplain()' /></div></div></form>"
	Else
	Response.Write"<div class='divtr1'>没有找到<span style='color:#F00'>"&Key&"</span> 记录！</div>"
	End If
	Response.Write"</div>"
Response.Write"<div class='mian' id='explain' style='display:none'><div class='top'>操作说明</div><div class='divtr2' style='padding:5px'><li>提升特别等级组：选中的会员提升后，不受帖数限制，并能享用一些论坛权限。</li><li>暂时删除：暂时删除选择的用户，只做标记，可以随时恢复！</li><li>完全删除：完全删除选择的用户，包括其帖子和留言等，删除后不可恢复！</li><li>删除帖子：删除选择的用户的所有帖子！</li><li>屏蔽帖子：屏蔽选择的用户的所有帖子，帖子将不能显示！</li><li>屏蔽签名：屏蔽选择的用户的签名后，帖子不会显示签名档！ </li><li>提升为VIP：直接将选择的用户提升为VIP会员！ </li><li>修复：整理修复选择用户的等级、总帖数、精华帖数等。</li></div></div>"
End Sub
%>