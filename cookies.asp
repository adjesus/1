<!-- #include File="Conn.asp" -->
<%
Dim GoUrl,Temp
GoUrl=Request.ServerVariables("HTTP_REFERER")
Temp=GetMemor("","CookiesDate")
If Int(Temp)>0 Then Response.Cookies(CacheName).Expires=date+Temp

Select Case Lcase(Request("action"))
Case"style"
	Style()
Case"font"
	Font
Case"online"
	MyOnline
End Select
If isnull(GoUrl) or GoUrl="" Then GoUrl="Index.asp"
Response.Redirect GoUrl

Sub Style()
	Response.Cookies(CacheName&"skinid").Expires= Date+7
	If Not Isnumeric(Request("skinid")) Then Exit Sub
	LetMemor "skinid","skinid",Request("skinid")
End Sub

Sub Font()
	Dim Size
	Size=Request("Size")
	If Size="9" or Size="10" or Size="12" Then LetMemor "","FontSize",Size
End Sub

Sub MyOnline
	'If GetMemor("","MyHidden")="1" Then
	If Request("ID")="0" Then
		LetMemor "","MyHidden",0
	Else
		LetMemor "","MyHidden",1
	End If
	Session(CacheName & "Stats")="BBS"
End Sub

Sub LetMemor(root,name,value)
	Session(CacheName & name)=value
	Response.Cookies(CacheName & root)(name)=value
End Sub
Function GetMemor(root,name)
	GetMemor=Request.Cookies(CacheName & root)(name)
	If GetMemor="" Then GetMemor=Session(CacheName & name)
End Function
%>