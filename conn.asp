<%@LANGUAGE="VBSCRIPT"%>
<%
Option Explicit
Response.Buffer = True
'Session.CodePage=936
Const Timeset=0 'ʱ����Զ�����(����ʱ��)
Dim Conn,StartTime,PageUrl,CacheName
StartTime = Timer()
PageURL=Lcase(Request.ServerVariables("URL"))
CacheName="BBS"&Replace(left(PageURL,instrRev(PageURL,"/")),"/","")
Sub ConnectionDatabase
	Dim Db,ConnStr
	on error resume next
	Db="mdb/344##4674@#.mdb"
	Set conn=Server.CreateObject("ADODB.Connection")
	ConnStr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&server.mappath(Db)
	Conn.Open ConnStr
	If Err Then
		Err.Clear
		Set Conn = Nothing
		Response.Write "�������Ӵ���!"
		Response.End
	End If
End Sub
%>