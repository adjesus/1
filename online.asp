<!--#include file="inc.asp"-->
<%
Dim ID,OnlineList,OnlinePic,Onlinedisplay
ID=request.querystring("ID")
If ID="1" Then Session(CacheName&"Online")=""
If Session(CacheName&"Online")="1" Then
	If ID<>"1" Then Session(CacheName&"Online")="0"
	OnlineList=""
	Onlinedisplay = "none"
	OnlinePic = BBS.ReadSkins("i@SkinDir")&"/+.gif"
Else
	Session(CacheName&"Online")="1"
	OnlineList=ShowOnlineList()
	Onlinedisplay = ""
	OnlinePic = BBS.ReadSkins("i@SkinDir")&"/-.gif"
End If
Set BBS =Nothing

Function ShowOnlineList()
Dim Temp,list,OnlineCache,AllonlineNum,EachOnline,User,S,I,II,pic,UserIP,PageInfo,TempBoard
Dim PSize,page,PageMax,Title
If BBS.Cache.valid("OnlineCache") Then
	OnlineCache=BBS.Cache.Value("OnlineCache")
	EachOnline=Split(OnlineCache,",")
	AllonlineNum=uBound(EachOnline)-1
	If BBS.BoardID<>0 Then
		For I=0 To AllonlineNum
			Temp=Split(EachOnline(i),"|")
			If Int(Temp(8))=BBS.BoardID Then
			TempBoard=TempBoard&EachOnline(i)&","
			End If
		Next
			OnlineCache=TempBoard
			EachOnline=Split(OnlineCache,",")
			AllonlineNum=uBound(EachOnline)-1
	End If
	PSize=Int(BBS.Info(47))
	page=Request("page")
	If not isnumeric(page) then Page=1
	page=int(page)
	If AllOnlineNum-1 mod PSize = 0 Then
		PageMax = AllOnlineNum \ PSize
	Else
		PageMax = AllOnlineNum \ PSize + 1
	End If
	If Page >PageMax Then Page=PageMax
	If Page<1 Then Page=1
	
	If AllonlineNum =>page*PSize Then AllonlineNum=page*PSize-1
	For i=(page*PSize-pSize) to AllonlineNum
	II=II+1
	Temp=Split(EachOnline(i),"|")
	User=Temp(1)
	UserIp="�����ñ���"
	Select Case Temp(6)
	Case "9"
	Pic=BBS.SkinsPic(21)
	Case "8"
	Pic=BBS.SkinsPic(22)
	Case "7"
	Pic=BBS.SkinsPic(23)
	Case "4"
	Pic=BBS.SkinsPic(24)
	Case "1"
	If BBS.MyAdmin<>9 Then User="��ʿ"
	Pic=BBS.SkinsPic(26)
	Case "0"
	If User="" Then
		Pic=BBS.SkinsPic(27)
		User="�ÿ�"
	Else
		Pic=BBS.SkinsPic(25)
	End If
	End Select
	If BBS.FoundUser Then
		If SESSION(CacheName& "MyGradeInfo")(42)="1" Then
			UserIP=Temp(5)
		End If
	End If
	Title="����λ�ã�"&Temp(7)&"&#10&#13����ʱ�䣺"&Temp(3)&"&#10&#13�ʱ�䣺"&Temp(4)&"&#10&#13��ʵIP��ַ��"&UserIp
	If User="�ÿ�" or User="��ʿ" Then
	    If Temp(1)<>"" And Temp(1)=BBS.MyName Then
		  User = " <span title='����������Լ�Ŷ' style='color:red'>"&User&"</span>"
		Else
		  User = " <a title="""&Title&""" >"&User&"</a>"
		End If
	Else
	    If Temp(1)<>"" And Temp(1)=BBS.MyName Then
		  User = " <span title='����������Լ�Ŷ' style='color:red'>"&User&"</span>"
		Else
		  User=" <a title='"&Title&"' href='userinfo.asp?name="&User&"'>"&User&"</a>"
		End If
	End If
	List=List&"<td width='10%'>"&pic&User&"</td>"
	If II mod 10 =0 And II<>PSize Then List=List&"</tr><tr>"
Next
	PageInfo="ҳ�Σ�"&Page&" / "&PageMax&"ҳ"
	if Page<>1 then
		PageInfo=PageInfo&"��<a target='hiddenframe' href='online.asp?page=1&boardid="&BBS.BoardID&"&id=1'>��ҳ</a>��"
		PageInfo=PageInfo& "<a target='hiddenframe' href='online.asp?page="&cstr(Page-1)&"&boardid="&BBS.BoardID&"&id=1'>����һҳ��</a>"
	end if
	If PageMax-Page>=1 then
		PageInfo=PageInfo& "<a target='hiddenframe' href='online.asp?page="&cstr(Page+1)&"&BoardID="&BBS.BoardID&"&id=1'>����һҳ��</a>"
		PageInfo=PageInfo& "<a target='hiddenframe' href='online.asp?page="&PageMax&"&BoardID="&BBS.BoardID&"&id=1'>��βҳ��</a>"
	End if
	List="<table border='0' width='100%'><tr>"&List&"</tr></table>"
	S=BBS.ReadSkins("��ʾ�����б�")
	S=Replace(S,"{����Ա}",BBS.SkinsPic(21))
	S=Replace(S,"{��������}",BBS.SkinsPic(22))
	S=Replace(S,"{����}",BBS.SkinsPic(23))
	S=Replace(S,"{VIP��Ա}",BBS.SkinsPic(24))
	S=Replace(S,"{��Ա}",BBS.SkinsPic(25))
	S=Replace(S,"{����}",BBS.SkinsPic(26))
	S=Replace(S,"{�ο�}",BBS.SkinsPic(27))
	S=Replace(S,"{�û��б�}",list)
	S=Replace(S,"{��ҳ}",PageInfo)
	S=Replace(S,CHR(34),CHR(39))
	S=Replace(S,VbCrlf,"")
	ShowOnlineList=S
End If
End Function
%>
<script language="JavaScript" type="text/JavaScript">
parent.document.getElementById("showon").style.display="<%=Onlinedisplay%>";
parent.document.getElementById("showon").innerHTML="<%=OnlineList%>";
parent.document.getElementById("onlinepic").src="Skins/<%=OnlinePic%>";
</script>