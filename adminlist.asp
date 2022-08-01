<!--#include file="Inc.asp"-->
<!--#include file="Inc/page_Cls.asp"-->
<%
Dim action,WhereSql,Title,Page_Url
If Request.QueryString("page") > 1 Then
  Page_Url = "?Page="&Request.QueryString("page")
Else
  Page_Url = ""
End If
BBS.Head"AdminList.asp"&Page_Url,"","查看管理团队"
action=Request.querystring("action")
Select Case action
Case "9"
	whereSql="BoardID=0"
	Title=BBS.GetGradeName(0,9)
Case "8"
	whereSql="BoardID=-1"
	Title=BBS.GetGradeName(0,8)
Case "7"
	whereSql="BoardID>0"
	Title=BBS.GetGradeName(0,7)
Case else
	whereSql=""
	Title="全部管理团队"
End select
ShowList()
BBS.Footer()
Set BBS =Nothing

Sub ShowList()
	Dim P,Page,arr_Rs,i,S,PInfo,Grade,BgColor
	Page = Request.QueryString("page")
	Set P = New Cls_PageView
	P.strTableName = "[Admin] As A inner join [User] As U on A.Name=U.Name"
	P.strPageUrl = "?action="&action
	P.strFieldsList = "A.Name,A.BoardID,U.Mail,U.lastTime,U.GradeID,U.GradeFlag"
	P.strCondiction = WhereSql
	P.strOrderList = "U.LastTime desc"
	P.strPrimaryKey = "A.Name"
	P.intPageSize = 20
	P.intPageNow = Page
	P.strCookiesName = "User_List"&action'客户端记录名称
	P.Reloadtime=10'分钟更新
	P.InitClass
	Arr_Rs = P.arrRecordInfo
	PInfo = P.strPageInfo
	Set P = nothing
	S="<div style='text-align:center;border-bottom:1px solid "&BBS.SkinsPIC(0)&";padding:3px;font-weight:bold;BACKGROUND: "&BBS.SkinsPIC(2)&";'><div style=""float:left;width:18%"">用户名称</div><div style=""float:left;width:18%"">论坛职位</div><div style=""float:left;width:20%"">管理区域</div><div style=""float:left;width:8%"">E-mail</div><div style=""float:left;width:8%"">留言</div><div style=""float:left;"">最后登陆</div><div style=""clear: both;""></div></div>"
	If IsArray(Arr_Rs) Then
	For i = 0 to UBound(Arr_Rs,2)
	If I mod 2<>0 Then Bgcolor="background:"&BBS.SkinsPIC(1) Else Bgcolor="" 
	S=S&"<div style='text-align:center;border-bottom:1px solid "&BBS.SkinsPIC(0)&";overflow:hidden;height:22px;line-height:22px;"&BgColor&"'><div style=""float:left;width:18%;""><a href='userinfo.asp?name="&Arr_Rs(0,i)&"'>"&Arr_Rs(0,i)&"</a></div><div style=""float:left;width:18%;border-left:1px solid "&BBS.SkinsPIC(0)&";""><a href=""help.asp?action=gradestring&id="&Arr_Rs(4,i)&""">"&BBS.GetGradeName(Arr_Rs(4,i),0)&"</a></div><div style=""float:left;width:20%;border-left:1px solid "&BBS.SkinsPIC(0)&";"">"&BBS.GetBoardName(Arr_Rs(1,i))&"</div><div style=""float:left;width:8%;height:22px;border-left:1px solid "&BBS.SkinsPIC(0)&";""><script language=""JavaScript"" type=""text/javascript"">mail('"&BBS.Fun.GetSqlStr(Arr_Rs(2,i))&"')</script></div><div style=""float:left;width:8%;height:22px;border-left:1px solid "&BBS.SkinsPIC(0)&";""><a href='sms.asp?action=write&name="&Arr_Rs(0,i)&"'><img src='images/Icon/sms.gif' border='0' /></a></div><div style=""float:left;border-left:1px solid "&BBS.SkinsPIC(0)&";padding-left:5px;"">"&Arr_Rs(3,i)&"</div><div style=""clear: both;""></div></div>"
	Next
	End If
	S=S&"<div style=""BACKGROUND: "&BBS.SkinsPIC(2)&";"">"&PInfo&"</div>"
	BBS.ShowTable "管理团队","<div style=""text-align:center;padding:3px""><a href=""?"">全部管理团队</a>&nbsp;|&nbsp;<a href=""?action=9"">"&BBS.GetGradeName(0,9)&"</a>&nbsp;|&nbsp;<a href=?action=8>"&BBS.GetGradeName(0,8)&"</a>&nbsp;|&nbsp;<a href=?action=7>"&BBS.GetGradeName(0,7)&"</a></div>"
	BBS.ShowTable Title,S
End Sub
%>
