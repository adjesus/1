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
Response.Write"<div class='mian'><div class='top'>�������йصı���</div>"
divBBS "ASP�ű���������",ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion,2
divBBS "��ȡISAPIDLL��metabase·��",request.ServerVariables("APPL_MD_PATH"),1
divBBS "��ʾվ������·��",request.ServerVariables("APPL_PHYSICAL_PATH"),2
divBBS "·����Ϣ",request.ServerVariables("PATH_INFO"),1
divBBS "��ʾ�������IP��ַ",request.ServerVariables("REMOTE_ADDR"),2
divBBS "������IP��ַ",Request.ServerVariables("LOCAL_ADDR"),1
divBBS "��ʾִ��SCRIPT������·��",request.ServerVariables("SCRIPT_NAME"),2
divBBS "���ط���������������DNS��������IP��ַ",request.ServerVariables("SERVER_NAME"),1
divBBS "���ط�������������Ķ˿�",request.ServerVariables("SERVER_PORT"),2
divBBS "Э������ƺͰ汾",request.ServerVariables("SERVER_PROTOCOL"),1
divBBS "�����������ƺͰ汾",request.ServerVariables("SERVER_SOFTWARE"),2
divBBS "����������ϵͳ",Request.ServerVariables("OS"),1
divBBS "�ű���ʱʱ��",Server.ScriptTimeout&"��",2
divBBS "������CPU����",Request.ServerVariables("NUMBER_OF_PROCESSORS")&"��",1
Response.Write"</div><div class='mian'><div class='top'>���֧�����</div>"
Response.Write"<div class='divtr2' style='padding:3px;'><form>�������֧�������ѯ��<INPUT size='30' name='classname'><INPUT type='submit' class='button' value='�� ѯ' />��������� ProgId �� ClassId </form></div>"
Dim strClass
strClass = Trim(Request("classname"))
If strClass<>"" then
Response.Write "<div class='divtr1' style='padding:3px;'>��ָ��������ļ������"
If Not IsObjInstalled(strClass) then 
Response.Write "<font color=red>���ź����÷�������֧��" & strclass & "�����</font>"
Else
Response.Write "<font color=green>��ϲ���÷�����֧��" & strclass & "�����</font>"
End If
Response.Write "</div>"
end if
Response.Write"</div><div class='mian'><div class='top'>IIS�Դ����</div>"
dim i,S,S1,Style
For i=0 to 10
If I mod 2=0 Then style=1 Else Style=2
	S=""
	select case i
	case 9
	S= "(FSO �ı��ļ���д)"
	case 10
	S= "(ACCESS ���ݿ�)"
	end select
	If Not IsObjInstalled(theInstalledObjects(i)) Then 
	S1= "<span style='color:#F11'><b>��</b></span>"
	Else
	S1="<b>��</b>"
	End If
divBBS theInstalledObjects(i)&S,S1,Style
Next
Response.Write"</div><div class='mian'><div class='top'>�����������</div>"
For i=11 to UBound(theInstalledObjects)
If I mod 2=0 Then style=1 Else Style=2
S=""
select case i
case 11
S= "(SA-FileUp �ļ��ϴ�)"
case 12
S= "(SA-FM �ļ�����)"
case 13
S="(JMail �ʼ�����)"
case 14
S="(CDONTS �ʼ����� SMTP Service)"
case 15
S="(ASPEmail �ʼ�����)"
case 16
S="(LyfUpload �ļ��ϴ�)"
case 17
S="(ASPUpload �ļ��ϴ�)"
end select
	If Not IsObjInstalled(theInstalledObjects(i)) Then 
		S1= "<span style='color:#F11'><b>��</b></span>"
	Else
		S1="<b>��</b>"
	End If
divBBS theInstalledObjects(i)&S,S1,Style
Next
Response.Write"</div>"
Response.Write"<div class='mian'><div class='top'>��ʾ�ͻ�����������HTTP����</div>"
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
