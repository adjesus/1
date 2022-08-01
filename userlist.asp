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
	BBS.Head "UserList.asp?action="&action&"&order="&order&Page_Url,"","今日到访会员"
	WhereSql="IsDel=0 And DATEDIFF('d',[LastTime],'"&BBS.NowBbsTime&"')<1"
Else
	BBS.Head "UserList.asp?action="&action&"&order="&order&Page_Url,"","查看用户列表"
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
	Response.Write"<div style='padding:5px'><div style='text-align:center;width:350px;float:right'>排序方式：<a href=?action="&action&">顺</a> / <a href=?action="&action&"&order=2>倒</a></div>查看按：<a href=?action=sex&order="&order&">姓别</a>&nbsp;/&nbsp;<a href=?action=essay&order="&order&">贴数</a>&nbsp;/&nbsp;<a href=?action=coin&order="&order&">金钱</a>&nbsp;/&nbsp;<a href=?action=mark&order="&order&">积分</a>&nbsp;/&nbsp;<a href=?action=regtime&order="&order&">注册时间</a>&nbsp;/&nbsp;<a href=?action=grade&order="&order&">等级</a>&nbsp;/&nbsp;<a href=?action=today>今日到访</a></div>"
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
	'P.CountSQL=200'为减轻服务器负担，只显示前200名
	P.strPrimaryKey = "ID"
	P.intPageSize = 40
	P.intPageNow = Page
	P.strCookiesName = "User_List"&action'客户端记录总数
	P.Reloadtime=10'每10分钟更新Cookies
	P.strPageVar = "page"
	P.InitClass
	Arr_Rs = P.arrRecordInfo
	PInfo = P.strPageInfo
	page=P.intPageNow
	Set P = nothing
	S="<div style='text-align:center;border-bottom:1px solid "&BBS.SkinsPIC(0)&";padding:3px;font-weight:bold;BACKGROUND:"&BBS.SkinsPIC(2)&";'><div style=""float:left;width:15%"">用户名(等级)</div><div style=""float:left;width:10%"">性别</div><div style=""float:left;width:10%"">E-mail</div><div style=""float:left;width:10%"">发帖数</div><div style=""float:left;width:10%"">金钱</div><div style=""float:left;width:10%"">积分</div><div style=""float:left;width:15%"">注册时间</div><div style=""clear: both;""></div></div>"
	If IsArray(Arr_Rs) Then
	For i = 0 to UBound(Arr_Rs, 2)
	If I mod 2=0 Then Bgcolor="background:"&BBS.SkinsPIC(1) Else Bgcolor="" 
	If Arr_Rs(1,i) Then Temp="帅哥" Else Temp="靓女"
	S=S&"<div style='text-align:center;border-bottom:1px solid "&BBS.SkinsPIC(0)&";padding:3px;'><div style=""float:left;width:15%""><a href='userinfo.asp?name="&Arr_Rs(0,i)&"'>"&Arr_Rs(0,i)&"</a>("&BBS.GetGradeName(Arr_Rs(7,i),0)&")</div><div style=""float:left;width:10%"">"&Temp&"</div><div style=""float:left;width:10%""><script language=""JavaScript"" type=""text/javascript"">mail('"&BBS.Fun.GetSqlStr(Arr_Rs(2,i))&"')</script></div><div style=""float:left;width:10%"">"&Arr_Rs(3,i)&"</div><div style=""float:left;width:10%"">"&Arr_Rs(4,i)&"</div><div style=""float:left;width:10%"">"&Arr_Rs(5,i)&"</div><div style=""float:left;width:15%"">"&Formatdatetime(Arr_Rs(6,i),1)&"</div><div style=""clear: both;""></div></div>"
	Next
	End If
	S=S&"<div style=""BACKGROUND: "&BBS.SkinsPIC(2)&";"">"&PInfo&"</div>"
	BBS.ShowTable "用户列表",S
End Sub
%>