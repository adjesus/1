<!--#include file="Admin_Check.asp"-->
<!--#include file="inc/page_Cls.asp"-->
<%
Dim TopicFile,Del
TopicFile=BBS.Info(36)&"/"
Del="UploadFile/Del/"'�ƶ��ļ���Ŀ¼
Head()
CheckString "35"
ShowTable "�ϴ��ļ�����","<center><a href=?>�����ϴ���¼</a> |  <a href='?Action=delnouse'>���������ϴ��ļ�</a> | <a href=?Action=delnovisit>����û�з��ʵ��ļ�</a> | <a href=?Action=deluphalfyear>���������ϴ��ļ�</a></center>"
Select Case Request("Action")
Case"deluphalfyear"
	deluphalfyear
Case"delnovisit"
	delnovisit
Case"delnouse"
	delnouse
Case"delall"
	DelAll
Case"DelOptFile"
	DelOptFile
Case Else
	UploadFile
end select
Footer()

Rem #���ĺ���(2005-5-27)
Function FileList(str)
	Dim re,Test,temp
	Dim LoopCount
	Set re=new RegExp
	re.IgnoreCase =True
	re.Global=True
	LoopCount=0
	Str = Replace(Str, chr(10), "")
	Do While True
		re.Pattern="\[upload=(.[^\[]*)\]"
		Test=re.Test(Str)
		If Test Then
			re.Pattern="\[\/upload\]"
			Test=re.Test(Str)
			If Test Then
				re.Pattern="(^.*)\[upload=(.[^\[]*)\](.[^\[]*)\[\/upload\](.*)"
				Temp=Temp&re.Replace(Str,"$3")&","
				Str=re.Replace(Str,"$1$4")
			Else
				Exit Do
			End If 
		Else
			Exit Do
		End If
		LoopCount=LoopCount + 1
		If LoopCount>40 Then Exit Do'��ֹ��ѭ��
	Loop
	Set re=nothing
	FileList=Temp
End Function

Sub UploadFile
	Dim strPageInfo,arr_Rs,i,P,FileType
	Response.Write"<form name='kk' method='POST' action='?Action=DelOptFile'>"
	Response.Write"<div class='mian'><div class='top'>�û��ļ��ϴ���¼</div>"
	Set P = New Cls_PageView
	P.strTableName = "[UpFile]"
	P.strFieldsList = "FileID,FileName,userName,FileType,FileSize,UpTime,Hits"
	P.strPrimaryKey = "FileID"
	P.strOrderList = "FileID desc"
	P.intPageNow = Request("page")
	P.intPageSize = 25
	P.strCookiesName = "UpFile"'cookies����
	P.InitClass
	Arr_Rs = P.arrRecordInfo
	strPageInfo = P.strPageInfo
	Set P = nothing
	If IsArray(Arr_Rs) Then
		Response.Write"<table class='Stable' border='0' cellpadding='3' cellspacing='0' ><tbody><tr><th width='5%'>ѡ��</th><th width='40%'>�ϴ����ļ�</th><th width='10%'>����</th><th width='15%'>�ϴ��û�</th><th width='18%'>�ϴ�����</th><th width='12%'>��С</th></tr>"  
		For i = 0 to UBound(Arr_Rs, 2)
		Response.Write"<tr>"
		Response.Write"<td align='center'><input type='checkbox' name='ID' value='"&Arr_rs(1,i)&"' /><td>"
		FileType=Lcase(Split(Arr_rs(1,i),".")(ubound(Split(Arr_rs(1,i),"."))))
		If Instr("|"&LCase(BBS.Info(34))&"|","|"&FileType&"|")>0 Then Response.Write"<div style='float:right;'>���أ�"&Arr_rs(6,i)&"��</div>"
		Response.Write"<a href='"&BBS.Info(36)&"/"&Arr_rs(1,i)&"' target='_blank'>"&Arr_rs(1,i)&"</a></td>"&_        
		"<td align='center'><img src='pic/FileType/"&Arr_rs(3,i)&".gif' /></td><td align='center'>"&Arr_rs(2,i)&"</td><td align='center'>"&Arr_rs(5,i)&"</td><td align='center'>"&Arr_rs(4,i)&"</td></tr>"
		Next
		Response.Write"</tbody></table><div class='bottom'><input type=checkbox name=chkall value=on onClick='CheckAll(this.form)'> ȫѡ&nbsp;&nbsp;<input class='button' value='ɾ����ѡ' type='button'  onclick=""if(confirm('ɾ���󽫲��ָܻ�����ȷ��Ҫɾ����'))form.submit()"" /></div><div class='divtr2'>"&strPageInfo&"</div>"
	Else
	Response.Write"<div class='bottom'>û���ϴ��ļ��ļ�¼</div>"
	End If
	Response.Write"</div></form>"
End Sub
'��ȡ��������
Sub Delnouse
Dim go
go=Request("go")
If go="ok" Then
	LoginTxt "���ڶ�ȡ����,ʱ����ܻ�ܳ�"
	Dim Alltable,i,temp
	AllTable=Split(BBS.BBStable(0),",")
	For i=0 To uBound(AllTable)
    Set Rs=BBS.Execute("Select Content From [Bbs"&AllTable(i)&"]")
	do while not rs.eof
	Temp=Temp&FileList(rs(0))
    rs.movenext
	loop
	rs.close
	Next
	ShowTable"������Ч�ϴ��ļ� �ڶ���","<form method=POST action='?Action=delall'><input name='files' type='hidden' value='"&temp&"'> ˵�����˲�����ɾ��û�������������ӵ������ļ���<br><input name='Go' type='radio' value='move' checked> �ƶ���<font color=red>UploadFile/Del/</font>Ŀ¼�У����飬Ϊ��ֹ��ɾ�����鿴�޴����ɾ�����Ŀ¼���ɣ�<br><input name='Go' type='radio' value='del'> ֱ�Ӵӿռ�ɾ�� <hr /><input value='ȷ ��' class='button' type='submit' /></form><script>document.getElementById('abc').style.display='none';</script>"
Else
	ShowTable"������Ч�ϴ��ļ� ��һ��","˵���������������û����ʾ�����ӵ������ϴ��ļ���<br />�˲��������ܴ������ķ�������Դ��������ʱ�ر���̳������ҹ����ʱ���С�<br />����ȡ�����벻Ҫˢ��������<hr /><li>��һ����<a href='?Action=delnouse&go=ok'>��ʼ���</a></li>"
End If
End Sub

'�������
Sub DelAll
	LoginTxt"���ڴ����ļ�"
	Dim Fso,Folder,Files,upname,bbsfiles,Go,S
	bbsFiles=Request.Form("files")
	Go=Request.Form("Go")
	If bbsFiles="" Then bbsFiles="0"
	Set Fso=server.createobject("scripting.filesystemobject")
	If not Fso.FolderExists(server.mappath(Del)) then Fso.CreateFolder(server.mappath(Del))
	Set Folder=fso.Getfolder(server.MapPath(TopicFile))
	Set files=folder.files
	For Each Upname In files
		If instr(LCase(bbsFiles),LCase(upname.name))<=0 then
		BBS.execute("Delete * From [UpFile] Where FileName='"&upname.name&"'")
		If Go="move" Then
			Fso.MoveFile Server.mappath(TopicFile&upname.name),server.mappath(Del&upname.name)
		Else
			Fso.DeleteFile(Server.MapPath(TopicFile&Upname.name))
		End If
		End If
	Next
	Set Folder=nothing
	Set Files=nothing
	Set Fso=nothing
	Response.Write"<script>document.getElementById('abc').style.display='none';</script>"
	If Go="move" Then
		S="���õ��ϴ��ļ��Ѿ���ת����"&Del&"Ŀ¼�� !"
	Else
		S="���õ��ϴ��ļ��Ѿ�ɾ�� !"
	End If
	BBS.NetLog "������̨_"&S
	Suc"",S,"?"
End Sub

'��������
Sub Deluphalfyear
	Dim Go,DelTime,Fso,Folder,Files,upname,S
	Go=Request.Form("Go")
	DelTime=Request.Form("DelTime")
	If Go="" And DelTime="" Then
		Response.Write"<form method='POST'>"
		ShowTable "���������������ǰ�ϴ����ļ�","<input name='Go' type='radio' value='move' checked /> �ƶ���<font color=red>"&Del&"</font>Ŀ¼�У�Ϊ��ֹ��ɾ�����鿴�޴����ɾ�����Ŀ¼���ɣ�<br><input name='Go' type='radio' value='del'> ֱ�Ӵӿռ�ɾ�� <hr>������<input name='DelTime' type='text' class='text' size='4' value='180'>����ǰ�ϴ����ļ� <input value=' ȷ �� ' type='submit' class='button'></form>"
	Else
		If Not isnumeric(DelTime) Then GoBack "","����������������д��" :Exit Sub
		LoginTxt "���ڴ����ļ�"
	Set Fso=server.createobject("scripting.filesystemobject")
	If not Fso.FolderExists(server.mappath(Del)) then Fso.CreateFolder(server.mappath(Del))
	Set Folder=fso.Getfolder(server.MapPath(TopicFile))
	Set Files=Folder.files
	For Each upName In Files
	BBS.execute("Delete * From [UpFile] Where FileName='"&upname.name&"'")
		If datediff("D",upName.datecreated,now)>DelTime then
		If Go="move" Then
			Fso.MoveFile Server.mappath(TopicFile&upname.name),server.mappath(Del&upname.name)
		Else
			Fso.DeleteFile(Server.MapPath(TopicFile&Upname.name))
		End If
		End if
	Next
	Set Folder=nothing
	Set Files=nothing
	Set Fso=nothing
	Response.Write"<script>document.getElementById('abc').style.display='none';</script>"
	If Go="move" Then
	S="��"&DelTime&"����ǰ�ϴ����ļ��Ѿ���ת����"&Del&"Ŀ¼�� !"
	Else
	S="��"&DelTime&"����ǰ�ϴ����ļ��Ѿ�ɾ����"
	End If
	BBS.NetLog "������̨_"&S
	Suc"",S,"?"
	End IF
End Sub

'����û�з��ʵ��ļ�
Sub DelNoVisit
	Dim Go,DelTime,Fso,Folder,Files,upname,S
	Go=Request.Form("Go")
	DelTime=Request.Form("DelTime")
	If Go="" And DelTime="" Then
	Response.Write"<form method='POST' style='margin:0px'>"
		ShowTable"�����������ǰû�з��ʵ��ϴ��ļ�","<input name='Go' type='radio' value='move' checked> �ƶ���<font color=red>"&Del&"</font>Ŀ¼�У�Ϊ��ֹ��ɾ�����鿴�޴����ɾ�����Ŀ¼���ɣ�<br><input name='Go' type='radio' value='del'> ֱ�Ӵӿռ�ɾ�� <hr>������<input name='DelTime' size=4 type='text' value='60'>����ǰû�з��ʵ��ϴ��ļ� <input value=' ȷ �� ' type=submit></form>"
	Else
		If Not isnumeric(DelTime) Then GoBack"","����������������д��":Exit Sub
		LoginTxt"���ڴ����ļ�"
		Set Fso=server.createobject("scripting.filesystemobject")
		If not Fso.FolderExists(server.mappath(Del)) then Fso.CreateFolder(server.mappath(Del))
		Set Folder=fso.Getfolder(server.MapPath(TopicFile))
		Set Files=Folder.files
		For Each Upname In Files
			if Datediff("d",UpName.DateLastAccessed,now)>DelTime then
			If Go="move" Then
				Fso.MoveFile Server.mappath(TopicFile&upname.name),server.mappath(Del&upname.name)
			Else
				Fso.DeleteFile(Server.MapPath(TopicFile&Upname.name))
			End If
			End if
		Next
		Set Folder=nothing
		Set Files=nothing
		Set Fso=nothing
		Response.Write"<script>document.getElementById('abc').style.display='none';</script>"
		If Go="move" Then
			S="����"&DelTime&"����ǰû�з��ʵ��ϴ��ļ��Ѿ���ת����"&Del&"Ŀ¼�� !"
		Else
			S="����"&DelTime&"����ǰû�з��ʵ��ϴ��ļ��Ѿ�ɾ�� !"
		End If
		BBS.NetLog "������̨_"&S
		Suc"",S,"?"
	End If
End Sub

'ɾ����ѡ
Sub DelOptFile
	Dim FileName,FSO,Folder,Files,Upname,Temp,i,S
	On Error Resume Next
    Set FSO = Server.CreateObject("Scripting.FileSystemObject")
		If Err Then
			Goback"","����ʧ�ܣ��ռ䲻֧��FOS�ļ���д����"
			err.Clear
			Exit Sub
		End If
	FileName=Request("ID")
	If FileName="" Then GoBack"","����ѡ����Ŀ��":Exit Sub
	Temp=Split(FileName,",")
	For i=0 To uBound(Temp)	
		BBS.execute("Delete * From [UpFile] Where FileName='"&Trim(Temp(i))&"'")
	Next
	Set Folder=fso.Getfolder(server.MapPath(BBS.Info(36)))
	Set files=folder.files
	For Each Upname In files
		If instr(LCase(FileName),LCase(Upname.name))>0 then
			FSO.DeleteFile(Server.MapPath(TopicFile&Upname.name))
		End if
	Next
	Set Folder=nothing
	Set Files=nothing
	Set Fso=nothing
	S="�ɹ�ɾ������ѡ���ϴ��ļ���"
	BBS.NetLog "������̨_"&S
	Suc"",S,"?"
End Sub

Sub LoginTxt(txt)
	Response.Write"<center><div id='abc' style='border:#999999 2px inset;margin:5px;background:#FFFF99;padding:10px;width:300px;color:#F00'><img src='Images/icon/await.gif'><br />"&Txt&"�����Ժ򡣡���</div></center>"
	Response.Flush
End Sub
%>
