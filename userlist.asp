<!--#include file="Inc.asp"-->
<!--#include file="Inc/page_Cls.asp"-->
<%
Dim action,WhereSql,orderSql,order,PageInfo,Title,Page_Url
If Not BBS.Founduser Then BBS.GoToErr(10)
order=Request.querystring("order")
action=Lcase(Request.querystring("action"))
If Request.QueryString("page") > 1 Then
  Page_Url = "&Page="&Request.QueryString("page")
Else
  Page_Url = ""
End If

If action="today" Then
	BBS.Head "UserList.asp?action="&action&"&order="&order&Page_Url,"","���յ��û�Ա"
	WhereSql="IsDel=0 And DATEDIFF('d',[LastTime],'"&BBS.NowBbsTime&"')<1"
Else
	BBS.Head "UserList.asp?action="&action&"&order="&order&Page_Url,"","�鿴�û��б�"
	WhereSql="IsDel=0"
End If
Select Case action
Case"sex"
	orderSql="Sex"
Case"coin"
	orderSql="Coin"
Case"essay"
	orderSql="EssayNum"
Case"mark"
	orderSql="Mark"
Case"regtime"
	orderSql="RegTime"
Case"grade"
	orderSql="GradeFlag"
Case else
	orderSql="ID"
End select
IF order="" Then orderSql =orderSql&" Desc"
ShowListTop()
ShowUserList()
BBS.Footer()
Set BBS =Nothing

Sub ShowListTop()
	Response.Write"<div style='padding:5px'><div style='text-align:center;width:350px;float:right'>����ʽ��<a href=?action="&action&">˳</a> / <a href=?action="&action&"&order=2>��</a></div>�鿴����<a href=?action=sex&order="&order&">�ձ�</a>&nbsp;/&nbsp;<a href=?action=essay&order="&order&">����</a>&nbsp;/&nbsp;<a href=?action=coin&order="&order&">��Ǯ</a>&nbsp;/&nbsp;<a href=?action=mark&order="&order&">����</a>&nbsp;/&nbsp;<a href=?action=regtime&order="&order&">ע��ʱ��</a>&nbsp;/&nbsp;<a href=?action=grade&order="&order&">�ȼ�</a>&nbsp;/&nbsp;<a href=?action=today>���յ���</a></div>"
End Sub
Sub ShowUserList()
	Dim P,Page,arr_Rs,i,Temp,S,PInfo,BgColor
	Page = Request.QueryString("page")
	Set P = New Cls_PageView
	P.strTableName = "[User]"
	P.strPageUrl = "?action="&action&"&order="&order
	P.strFieldsList = "Name,Sex,Mail,EssayNum,Coin,Mark,RegTime,GradeID,GradeFlag"
	P.strCondiction = WhereSql
	P.strorderList = orderSql
	'P.CountSQL=200'Ϊ���������������ֻ��ʾǰ200��
	P.strPrimaryKey = "ID"
	P.intPageSize = 40
	P.intPageNow = Page
	P.strCookiesName = "User_List"&action'�ͻ��˼�¼����
	P.Reloadtime=10'ÿ10���Ӹ���Cookies
	P.strPageVar = "page"
	P.InitClass
	Arr_Rs = P.arrRecordInfo
	PInfo = P.strPageInfo
	page=P.intPageNow
	Set P = nothing
	S="<div style='text-align:center;border-bottom:1px solid "&BBS.SkinsPIC(0)&";padding:3px;font-weight:bold;BACKGROUND:"&BBS.SkinsPIC(2)&";'><div style=""float:left;width:15%"">�û���(�ȼ�)</div><div style=""float:left;width:10%"">�Ա�</div><div style=""float:left;width:10%"">E-mail</div><div style=""float:left;width:10%"">������</div><div style=""float:left;width:10%"">��Ǯ</div><div style=""float:left;width:10%"">����</div><div style=""float:left;width:15%"">ע��ʱ��</div><div style=""clear: both;""></div></div>"
	If IsArray(Arr_Rs) Then
	For i = 0 to UBound(Arr_Rs, 2)
	If I mod 2=0 Then Bgcolor="background:"&BBS.SkinsPIC(1) Else Bgcolor="" 
	If Arr_Rs(1,i) Then Temp="˧��" Else Temp="��Ů"
	S=S&"<div style='text-align:center;border-bottom:1px solid "&BBS.SkinsPIC(0)&";padding:3px;'><div style=""float:left;width:15%""><a href='userinfo.asp?name="&Arr_Rs(0,i)&"'>"&Arr_Rs(0,i)&"</a>("&BBS.GetGradeName(Arr_Rs(7,i),0)&")</div><div style=""float:left;width:10%"">"&Temp&"</div><div style=""float:left;width:10%""><script language=""JavaScript"" type=""text/javascript"">mail('"&BBS.Fun.GetSqlStr(Arr_Rs(2,i))&"')</script></div><div style=""float:left;width:10%"">"&Arr_Rs(3,i)&"</div><div style=""float:left;width:10%"">"&Arr_Rs(4,i)&"</div><div style=""float:left;width:10%"">"&Arr_Rs(5,i)&"</div><div style=""float:left;width:15%"">"&Formatdatetime(Arr_Rs(6,i),1)&"</div><div style=""clear: both;""></div></div>"
	Next
	End If
	S=S&"<div style=""BACKGROUND: "&BBS.SkinsPIC(2)&";"">"&PInfo&"</div>"
	BBS.ShowTable "�û��б�",S
End Sub
%>