<!-- #include File="Inc.asp" -->
<!-- #include File="Inc/Upload_Cls.asp" -->
<%
Server.ScriptTimeOut=300
dim FilePath,FacePath
FilePath = BBS.Info(36)
FacePath = BBS.Info(37)
Dim Upload,ReturnString,Temp,Flag
BBS.CheckMake'��ֹ�ⲿ�ύ
Flag = Request.QueryString("Flag")
If Not BBS.FoundUser Then BBS.GoToErr(10)
Set Upload = New Upload_Cls
If BBS.Info(30)="0" Then upload.ErrPrint"����̳���ý�ֹ�ϴ�"
If Flag="0" Then
	If SESSION(CacheName& "MyGradeInfo")(14)="0" Then upload.ErrPrint"��û���ϴ��ĵȼ�Ȩ��" 
Else
	If SESSION(CacheName& "MyGradeInfo")(17)="0" Then upload.ErrPrint"��û���ϴ�ͷ��ĵȼ�Ȩ��"
End If 

Temp=BBS.Execute("Select Count(*) From[UpFile] where UserName='"&BBS.MyName&"' And DATEDIFF('d',[UpTime],'"&BBS.NowBbsTime&"')<1")(0)
If IsNull(Temp) Then Temp=SESSION(CacheName& "MyGradeInfo")(15)
If Int(Temp) => Int(SESSION(CacheName& "MyGradeInfo")(15)) Then Upload.MaxFile=0
Upload.FileTypeFlag = Replace(BBS.Info(34)&"|"&BBS.Info(35),"||","")
If Flag = "0" Then
	Upload.SaveData FilePath,"",0
	ReturnString = "<br>[UPLOAD=" & Upload.FileTypeName & "," & Upload.FileSizeKB & "," & Upload.ReWidth & "," & Upload.Width & "," & Upload.Height & "]" & Upload.FileName & "[/UPLOAD]"
	BBS.Execute("insert into [UpFile](FileName,FileType,FileSize,UpTime,UserName) values ('"&Upload.FileName & "','" & upload.FileTypeName & "'," & upload.FileSize & ",'"& BBS.NowBBSTime &"','" & BBS.MyName & "')")
	Response.Write("<body leftmargin=""0"" topmargin=""0"" onload=""javascript:parent.IframeID.document.body.innerHTML+='"&ReturnString&"';"">")
Else
	Upload.SaveData FacePath,BBS.MyID,1
	ReturnString =  "viewfile.asp?Path=Face&FileName=" & Upload.FileName
	Response.Write("<body leftmargin=""0"" topmargin=""0"" onload=""javascript:parent.document.getElementById('pic').src='"&ReturnString&"';parent.document.getElementById('picurl').value='"&ReturnString&"';parent.document.getElementById('PicW').value='"&Upload.Width&"';parent.document.getElementById('PicH').value='"&Upload.Height&"';"">")
End If
Upload.ErrPrint "�ϴ��ɹ�"
Set Upload=Nothing
Set BBS =Nothing
%>

