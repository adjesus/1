<!--#include file="Inc.asp"-->
<%
'BBS2.0
Dim TFilePath
On Error Resume Next
Response.Buffer = True
Response.Clear
If BBS.Info(31)="1" Then ForbidMake
If LCase(Request("Path"))="face" Then
   TFilePath=  BBS.Info(37) & "/" & Request("FileName")
Else
   TFilePath = BBS.Info(36) & "/" & Request("FileName")
End If

If ChkFile(TFilePath) Then
	Response.Redirect("Images/NoImg.gif")
End If

If BBS.Info(38)="1" Then
	DownloadFile(TFilePath)
Else
	Response.Redirect(TFilePath)
End If
Set BBS =Nothing

Sub ForbidMake()
	Dim Come,Here
	Come=Cstr(Request.ServerVariables("HTTP_REFERER"))
	Here=Cstr(Request.ServerVariables("SERVER_NAME"))
	If Come<>"" And Mid(Come,8,Len(Here)) <> Here Then
		Response.Redirect "Images/url.gif"
		Response.end
	End If		
End Sub

Function ChkFile(FileName)
	Dim Temp,FileType,F
	ChkFile=false
	FileType=Lcase(Split(FileName,".")(ubound(Split(FileName,"."))))
	Temp="|asp|aspx|cgi|php|cdx|cer|asa|"
	If Instr(Temp,"|"&FileType&"|")>0 Then ChkFile=True
	F=Replace(Request("FileName"),".","")
	If instr(1,F,chr(39))>0 or instr(1,F,chr(34))>0 or instr(1,F,chr(59))>0 then ChkFile=True

	If Not ChkFile Then
		If Instr("|"&LCase(BBS.Info(34)&"|"&BBS.Info(35))&"|","|"&FileType&"|")>0 Then
			If BBS.Info(32)="1" Then BBS.Execute("update [upFile] Set hits=hits+1 where FileName='"&Request("FileName")&"'")
		End If
	End If
End Function

Function DownloadFile(FileName)
Dim TempFileName,Fsize
    On error resume next
	Server.ScriptTimeOut=999999
	Response.Clear
    Dim FileType,ADS,StrFileName,Data
    FileType=Lcase(Split(FileName,".")(ubound(Split(FileName,"."))))
	StrFileName=Server.Mappath(FileName)
	TempFileName = Split(StrFileName,"\")(Ubound(Split(StrFileName,"\")))
    Set ADS = Server.CreateObject("ADODB.Stream") 
	ADS.Open
	ADS.Type = 1 
    ADS.LoadFromFile(StrFileName)
	Data=ADS.Read
	Fsize=Clng(lenb(Data))
	If Err Then
  	   Response.Redirect("Images/NoImg.gif")
	   'Response.Write("<h1>´íÎó: </h1>" & err.Description & "<p>")
       Response.End 
    End If
	ADS.Close
    If Response.IsClientConnected Then 
       If FileType="gif" Or FileType="jpg" Or FileType="jpeg" Or FileType="bmp" Then 
	      Response.ContentType = "image/*"
	   Else
	      Response.AddHeader "Content-Disposition", "attachment; filename=" & TempFileName
		  Response.ContentType = "application/ms-download"
	   End If
	   Response.AddHeader "Content-Length", Fsize
 	   Response.CharSet = "UTF-8" 
	   Response.ContentType = "application/octet-stream" 
	   Response.BinaryWrite Data
	   Response.Flush
	End If
End Function
%>