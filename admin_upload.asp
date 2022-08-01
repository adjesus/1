<!--#include file="Admin_Check.asp"-->
<!--#include file="inc/page_Cls.asp"-->
<%
Dim TopicFile,Del
TopicFile=BBS.Info(36)&"/"
Del="UploadFile/Del/"'移动文件的目录
Head()
CheckString "35"
ShowTable "上传文件管理","<center><a href=?>管理上传记录</a> |  <a href='?Action=delnouse'>清理无用上传文件</a> | <a href=?Action=delnovisit>清理没有访问的文件</a> | <a href=?Action=deluphalfyear>批量清理上传文件</a></center>"
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

Rem #核心函数(2005-5-27)
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
		If LoopCount>40 Then Exit Do'防止死循环
	Loop
	Set re=nothing
	FileList=Temp
End Function

Sub UploadFile
	Dim strPageInfo,arr_Rs,i,P,FileType
	Response.Write"<form name='kk' method='POST' action='?Action=DelOptFile'>"
	Response.Write"<div class='mian'><div class='top'>用户文件上传记录</div>"
	Set P = New Cls_PageView
	P.strTableName = "[UpFile]"
	P.strFieldsList = "FileID,FileName,userName,FileType,FileSize,UpTime,Hits"
	P.strPrimaryKey = "FileID"
	P.strOrderList = "FileID desc"
	P.intPageNow = Request("page")
	P.intPageSize = 25
	P.strCookiesName = "UpFile"'cookies名称
	P.InitClass
	Arr_Rs = P.arrRecordInfo
	strPageInfo = P.strPageInfo
	Set P = nothing
	If IsArray(Arr_Rs) Then
		Response.Write"<table class='Stable' border='0' cellpadding='3' cellspacing='0' ><tbody><tr><th width='5%'>选择</th><th width='40%'>上传的文件</th><th width='10%'>类型</th><th width='15%'>上传用户</th><th width='18%'>上传日期</th><th width='12%'>大小</th></tr>"  
		For i = 0 to UBound(Arr_Rs, 2)
		Response.Write"<tr>"
		Response.Write"<td align='center'><input type='checkbox' name='ID' value='"&Arr_rs(1,i)&"' /><td>"
		FileType=Lcase(Split(Arr_rs(1,i),".")(ubound(Split(Arr_rs(1,i),"."))))
		If Instr("|"&LCase(BBS.Info(34))&"|","|"&FileType&"|")>0 Then Response.Write"<div style='float:right;'>下载："&Arr_rs(6,i)&"次</div>"
		Response.Write"<a href='"&BBS.Info(36)&"/"&Arr_rs(1,i)&"' target='_blank'>"&Arr_rs(1,i)&"</a></td>"&_        
		"<td align='center'><img src='pic/FileType/"&Arr_rs(3,i)&".gif' /></td><td align='center'>"&Arr_rs(2,i)&"</td><td align='center'>"&Arr_rs(5,i)&"</td><td align='center'>"&Arr_rs(4,i)&"</td></tr>"
		Next
		Response.Write"</tbody></table><div class='bottom'><input type=checkbox name=chkall value=on onClick='CheckAll(this.form)'> 全选&nbsp;&nbsp;<input class='button' value='删除所选' type='button'  onclick=""if(confirm('删除后将不能恢复！您确定要删除吗？'))form.submit()"" /></div><div class='divtr2'>"&strPageInfo&"</div>"
	Else
	Response.Write"<div class='bottom'>没有上传文件的记录</div>"
	End If
	Response.Write"</div></form>"
End Sub
'记取帖子数据
Sub Delnouse
Dim go
go=Request("go")
If go="ok" Then
	LoginTxt "正在读取数据,时间可能会很长"
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
	ShowTable"清理无效上传文件 第二步","<form method=POST action='?Action=delall'><input name='files' type='hidden' value='"&temp&"'> 说明：此操作将删除没有在帖子上连接的无用文件。<br><input name='Go' type='radio' value='move' checked> 移动到<font color=red>UploadFile/Del/</font>目录中（建议，为防止误删除，查看无错后再删除这个目录即可）<br><input name='Go' type='radio' value='del'> 直接从空间删除 <hr /><input value='确 定' class='button' type='submit' /></form><script>document.getElementById('abc').style.display='none';</script>"
Else
	ShowTable"清理无效上传文件 第一步","说明：检测在帖子上没有显示或连接的无用上传文件。<br />此操作将可能大量消耗服务器资源，建议暂时关闭论坛或在深夜人少时进行。<br />检测读取过程请不要刷屏或点击。<hr /><li>第一步：<a href='?Action=delnouse&go=ok'>开始检测</a></li>"
End If
End Sub

'清除无用
Sub DelAll
	LoginTxt"正在处理文件"
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
		S="无用的上传文件已经被转移至"&Del&"目录下 !"
	Else
		S="无用的上传文件已经删除 !"
	End If
	BBS.NetLog "操作后台_"&S
	Suc"",S,"?"
End Sub

'批量清理
Sub Deluphalfyear
	Dim Go,DelTime,Fso,Folder,Files,upname,S
	Go=Request.Form("Go")
	DelTime=Request.Form("DelTime")
	If Go="" And DelTime="" Then
		Response.Write"<form method='POST'>"
		ShowTable "批量清理多少天以前上传的文件","<input name='Go' type='radio' value='move' checked /> 移动到<font color=red>"&Del&"</font>目录中（为防止误删除，查看无错后再删除这个目录即可）<br><input name='Go' type='radio' value='del'> 直接从空间删除 <hr>清理在<input name='DelTime' type='text' class='text' size='4' value='180'>天以前上传的文件 <input value=' 确 定 ' type='submit' class='button'></form>"
	Else
		If Not isnumeric(DelTime) Then GoBack "","天数必需用数字填写！" :Exit Sub
		LoginTxt "正在处理文件"
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
	S="在"&DelTime&"天以前上传的文件已经被转移至"&Del&"目录下 !"
	Else
	S="在"&DelTime&"天以前上传的文件已经删除！"
	End If
	BBS.NetLog "操作后台_"&S
	Suc"",S,"?"
	End IF
End Sub

'清理没有访问的文件
Sub DelNoVisit
	Dim Go,DelTime,Fso,Folder,Files,upname,S
	Go=Request.Form("Go")
	DelTime=Request.Form("DelTime")
	If Go="" And DelTime="" Then
	Response.Write"<form method='POST' style='margin:0px'>"
		ShowTable"清理多少天以前没有访问的上传文件","<input name='Go' type='radio' value='move' checked> 移动到<font color=red>"&Del&"</font>目录中（为防止误删除，查看无错后再删除这个目录即可）<br><input name='Go' type='radio' value='del'> 直接从空间删除 <hr>清理在<input name='DelTime' size=4 type='text' value='60'>天以前没有访问的上传文件 <input value=' 确 定 ' type=submit></form>"
	Else
		If Not isnumeric(DelTime) Then GoBack"","天数必需用数字填写！":Exit Sub
		LoginTxt"正在处理文件"
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
			S="超过"&DelTime&"天以前没有访问的上传文件已经被转移至"&Del&"目录下 !"
		Else
			S="超过"&DelTime&"天以前没有访问的上传文件已经删除 !"
		End If
		BBS.NetLog "操作后台_"&S
		Suc"",S,"?"
	End If
End Sub

'删除所选
Sub DelOptFile
	Dim FileName,FSO,Folder,Files,Upname,Temp,i,S
	On Error Resume Next
    Set FSO = Server.CreateObject("Scripting.FileSystemObject")
		If Err Then
			Goback"","操作失败，空间不支持FOS文件读写！。"
			err.Clear
			Exit Sub
		End If
	FileName=Request("ID")
	If FileName="" Then GoBack"","请先选择项目。":Exit Sub
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
	S="成功删除了所选的上传文件。"
	BBS.NetLog "操作后台_"&S
	Suc"",S,"?"
End Sub

Sub LoginTxt(txt)
	Response.Write"<center><div id='abc' style='border:#999999 2px inset;margin:5px;background:#FFFF99;padding:10px;width:300px;color:#F00'><img src='Images/icon/await.gif'><br />"&Txt&"，请稍候。。。</div></center>"
	Response.Flush
End Sub
%>
