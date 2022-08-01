<!--#include file="Admin_check.asp"-->
<%
Head()
CheckString "58"
Dim theInstalledObjects(17)
theInstalledObjects(0) ="MSWC.AdRotator"
theInstalledObjects(1) ="MSWC.BrowserType"
theInstalledObjects(2) ="MSWC.NextLink"
theInstalledObjects(3) ="MSWC.Tools"
theInstalledObjects(4) ="MSWC.Status"
theInstalledObjects(5) ="MSWC.Counters"
theInstalledObjects(6) ="IISSample.ContentRotator"
theInstalledObjects(7) ="IISSample.PageCounter"
theInstalledObjects(8) ="MSWC.PermissionChecker"
theInstalledObjects(9) ="Scripting.FileSystemObject"
theInstalledObjects(10) ="adodb.connection"
theInstalledObjects(11) ="SoftArtisans.FileUp"
theInstalledObjects(12) ="SoftArtisans.FileManager"
theInstalledObjects(13) ="JMail.SMTPMail"
theInstalledObjects(14) ="CDONTS.NewMail"
theInstalledObjects(15) ="Persits.MailSender"
theInstalledObjects(16) ="LyfUpload.UploadFile"
theInstalledObjects(17) ="Persits.Upload.1"
call servervar()
Footer()
Sub divBBS(S,S1,Style)
Response.Write"<div class='divtr"&Style&"' style='padding:3px;'><div style='float:right;width:50%'>"&S1&"</div>"&S&"</div>"
End Sub

sub servervar()
Response.Write"<div class='mian'><div class='top'>服务器有关的变量</div>"
divBBS "ASP脚本解译引擎",ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion,2
divBBS "检取ISAPIDLL的metabase路径",request.ServerVariables("APPL_MD_PATH"),1
divBBS "显示站点物理路径",request.ServerVariables("APPL_PHYSICAL_PATH"),2
divBBS "路径信息",request.ServerVariables("PATH_INFO"),1
divBBS "显示请求机器IP地址",request.ServerVariables("REMOTE_ADDR"),2
divBBS "服务器IP地址",Request.ServerVariables("LOCAL_ADDR"),1
divBBS "显示执行SCRIPT的虚拟路径",request.ServerVariables("SCRIPT_NAME"),2
divBBS "返回服务器的主机名，DNS别名，或IP地址",request.ServerVariables("SERVER_NAME"),1
divBBS "返回服务器处理请求的端口",request.ServerVariables("SERVER_PORT"),2
divBBS "协议的名称和版本",request.ServerVariables("SERVER_PROTOCOL"),1
divBBS "服务器的名称和版本",request.ServerVariables("SERVER_SOFTWARE"),2
divBBS "服务器操作系统",Request.ServerVariables("OS"),1
divBBS "脚本超时时间",Server.ScriptTimeout&"秒",2
divBBS "服务器CPU数量",Request.ServerVariables("NUMBER_OF_PROCESSORS")&"个",1
Response.Write"</div><div class='mian'><div class='top'>组件支持情况</div>"
Response.Write"<div class='divtr2' style='padding:3px;'><form>其它组件支持情况查询：<INPUT size='30' name='classname'><INPUT type='submit' class='button' value='查 询' />输入组件的 ProgId 或 ClassId </form></div>"
Dim strClass
strClass = Trim(Request("classname"))
If strClass<>"" then
Response.Write "<div class='divtr1' style='padding:3px;'>您指定的组件的检查结果："
If Not IsObjInstalled(strClass) then 
Response.Write "<font color=red>很遗憾，该服务器不支持" & strclass & "组件！</font>"
Else
Response.Write "<font color=green>恭喜！该服务器支持" & strclass & "组件。</font>"
End If
Response.Write "</div>"
end if
Response.Write"</div><div class='mian'><div class='top'>IIS自带组件</div>"
dim i,S,S1,Style
For i=0 to 10
If I mod 2=0 Then style=1 Else Style=2
	S=""
	select case i
	case 9
	S= "(FSO 文本文件读写)"
	case 10
	S= "(ACCESS 数据库)"
	end select
	If Not IsObjInstalled(theInstalledObjects(i)) Then 
	S1= "<span style='color:#F11'><b>×</b></span>"
	Else
	S1="<b>√</b>"
	End If
divBBS theInstalledObjects(i)&S,S1,Style
Next
Response.Write"</div><div class='mian'><div class='top'>其他常见组件</div>"
For i=11 to UBound(theInstalledObjects)
If I mod 2=0 Then style=1 Else Style=2
S=""
select case i
case 11
S= "(SA-FileUp 文件上传)"
case 12
S= "(SA-FM 文件管理)"
case 13
S="(JMail 邮件发送)"
case 14
S="(CDONTS 邮件发送 SMTP Service)"
case 15
S="(ASPEmail 邮件发送)"
case 16
S="(LyfUpload 文件上传)"
case 17
S="(ASPUpload 文件上传)"
end select
	If Not IsObjInstalled(theInstalledObjects(i)) Then 
		S1= "<span style='color:#F11'><b>×</b></span>"
	Else
		S1="<b>√</b>"
	End If
divBBS theInstalledObjects(i)&S,S1,Style
Next
Response.Write"</div>"
Response.Write"<div class='mian'><div class='top'>显示客户发出的所有HTTP标题</div>"
Response.Write"<div class='divtr1' style='padding:3px;'>"&request.ServerVariables("All_Http")&"</div>"
Response.Write"</div>"
end sub

Function IsObjInstalled(strClassString)
On Error Resume Next
IsObjInstalled = False
Err = 0
Dim xTestObj
Set xTestObj = Server.CreateObject(strClassString)
If 0 = Err Then IsObjInstalled = True
Set xTestObj = Nothing
Err = 0
End Function
%>
