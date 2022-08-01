<!-- #include File="Conn.asp" -->
<!-- #include File="Inc/Main_Cls.asp" -->
<!-- #include File="Inc/Fun_Cls.asp" -->
<%
Dim BBS
Set BBS = New Cls_jybbs
BBS.Config()
'If Instr(BBS.BbsURL,"admin_")=0 Then 
BBS.CheckUser()
%>