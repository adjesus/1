<!--#include file="inc.asp"-->
<!--#include file="inc/Cls_Brower.asp"-->
<%
Dim PageString
BBS.Head "index.asp","","��̳��ҳ"
If BBS.Info(20)="1" then ShowInfo()
ShowNewHot()
ShowBoard()
If BBS.Info(22)="1" then ShowBirthday()
If BBS.Info(23)="1" then ShowLink()
If BBS.Info(21)="1" then PageString=PageString&ShowOnline()
Response.Write PageString 
Response.Write "<iframe name='hiddenframe' frameborder='0'  height='0' id='hiddenframe'></iframe>"
If Session(CacheName&"online")="1" then Response.Write "<iframe frameborder='0' src='online.asp?id=1' height='0'></iframe>"
BBS.Footer()
Set BBS = Nothing

Sub ShowInfo()
	With BBS
	Dim S,OnlingType
	If .FoundUser Then
		S = .ReadSkins("�û���Ϣ")
		If Session(CacheName & "MyInfo")(11)="1" Then
			S=Replace(S,"{�û�ͷ��}","<img src='http://qqshow-user.tencent.com/"&Session(CacheName & "MyInfo")(10)&"/11/'>")
		Else
			S=Replace(S,"{�û�ͷ��}","<img src="&Session(CacheName & "MyInfo")(12)&" width="&Session(CacheName & "MyInfo")(13)&" height="&Session(CacheName & "MyInfo")(14)&" >")
		End if
		If .MyHidden="0" Then
			OnlingType="������"
		Else
			OnlingType="������"
		End If
		S=Replace(S,"{�û�����}","<a href='userinfo.asp' title='���Լ���������Ϣ'>"&.MyName&"</a>")
		S=Replace(S,"{����״̬}",OnlingType)
		S=Replace(S,"{����}",Session(CacheName & "MyInfo")(4))
		S=Replace(S,"{������}",Session(CacheName & "MyInfo")(6))
		S=Replace(S,"{��Ǯ��}",Session(CacheName & "MyInfo")(7))
		S=Replace(S,"{�ȼ�}",Session(CacheName & "MyGradeInfo")(2))
	Else
		S = .ReadSkins("�ο���Ϣ")
	End If 
	S=Replace(S,"{������}",.InfoUpdate(0))
	S=Replace(S,"{������}",.InfoUpdate(1))
	S=Replace(S,"{��������}",.InfoUpdate(2))
	S=Replace(S,"{��������}",.InfoUpdate(3))
	S=Replace(S,"{���������}",.InfoUpdate(4))
	S=Replace(S,"{��Ա��}",.InfoUpdate(5))
	S=Replace(S,"{�»�Ա����}",.InfoUpdate(6))
	If .Info(14)="1" Then
		S=Replace(S,"{��֤��}",BBS.GetiCode)
	Else
		S=Replace(S,"{��֤��}","")
	End If
	S=Replace(S,"{����}",.Placard(0))
	Response.Write S
	End With
End Sub

Sub ShowBoard()
	Dim Board,Rs,i,BigBoard,BoardChild,BoardString,BoardStr,II
	Dim area,area2,Child,width
	With BBS
	If Not IsArray(.Board_Rs) Then .GetBoardCache()
	If Not IsArray(.Board_Rs) Then Exit Sub
	area=.ReadSkins("���������")
	area2=.ReadSkins("�����������")
	For i=0 To Ubound(.Board_Rs,2)
'	ֻ��ʾ2��
	  If .Board_Rs(0,i)<2 Then
'		Board_Rs()=0���+1ID+2��ID+3����+4ͼƬ+5���+6����+7��֤��+8����̳����+9������+10����+11����
		If .Board_Rs(0,i)=0 Then
			BoardString=Split(.Board_Rs(11,i),"|")
			If BoardString(1)="1" Then
				BigBoard=area2
				Child=Int(.Board_Rs(8,i))
				If Child<Int(BoardString(2)) and Child>0 Then
					Width=100\Child
				Else
					width=100\Int(BoardString(2))
				End If
				II=0
			Else
				BigBoard=area
			End If
			If i >= 1 Then
				BoardStr=Replace(BoardStr,"{��ʾ���}",BoardChild)
				Board=Board&BoardStr
				BoardChild =""
			End IF
			BoardStr=Replace(BigBoard,"{��������}",.Board_Rs(3,i))
			BoardStr=Replace(BoardStr,"{����ID}",.Board_Rs(1,i))
			Else
				If BoardString(1)="1" or Session(CacheName& "BoardStyle")="1" Then
					Child=Child-Int(.Board_Rs(8,i))
					II=II+1
					If II=Int(BoardString(2)) or II=Child Then
					    BoardChild=BoardChild&"<div style='float:left;max-width:"&width&"%'>"&.GetBoardInfo("1",i)&"</div>"
					Else
					    BoardChild=BoardChild&"<div style='float:left;width:"&width&"%'>"&.GetBoardInfo("1",i)&"</div>"
					End If
				Else
					BoardChild=BoardChild&.GetBoardInfo("0",i)
				End IF
			End If
		End If
	Next
	BoardStr=Replace(BoardStr,"{��ʾ���}",BoardChild)
	Board = Board&BoardStr
	Response.Write Board
	End With
End Sub

Sub ShowBirthday()
	Dim S,STemp
    STemp=BBS.ReadSkins("��Ա����")
	If BBS.Cache.valid("Birthday") then
		S=Split(BBS.Cache.Value("Birthday"),"|")
		If S(0)="0" Then Exit Sub
	    STemp=Replace(STemp,"{�������ջ�Ա��}",S(0))
	    STemp=Replace(STemp,"{����}",S(1))
	    Response.Write STemp
	Else
		Dim Rs,Arr_Rs,I,Num,UserBirthday
		Set Rs=BBS.Execute("Select Name,Birthday From [User] where Month(Birthday)=Month(now) and day(Birthday)=day(now)")
		IF Not Rs.eof Then Arr_Rs=Rs.getrows()
		Rs.Close
		Set Rs=Nothing
		Num=0
		If IsArray(Arr_Rs) Then
			For i = 0 to UBound(Arr_Rs,2)
				Num=Num+1
				UserBirthday=UserBirthday&"ף <a href=""userinfo.asp?name="&Arr_Rs(0,i)&"""><font color=""#800000"">"&Arr_Rs(0,i)&"</font></a> ���տ���&nbsp;&nbsp;"
			Next
		End If
		UserBirthday = ""&UserBirthday&""
		If i>5 Then UserBirthday="<marquee width=""99%"" onMouseOver=""this.stop()"" onMouseOut=""this.start()"" scrollamount=""3"" >"&UserBirthday&"</marquee>"
		If Num=0 Then
			S="0|0"
		Else	 
	        STemp=Replace(STemp,"{�������ջ�Ա��}",num)
	        STemp=Replace(STemp,"{����}",UserBirthday)
	        Response.Write STemp
			S=Num&"|"&UserBirthday
		End If
		BBS.Cache.add "Birthday",S,dateadd("n",1200,now)
	End If
End Sub

Sub showlink()
	Dim rs,Arr_Rs,I,Temp,lpic,TempText,TempPic,ii,S
	If BBS.Cache.valid("linkinfo") then
		Temp=BBS.Cache.Value("linkinfo")
	Else	
		Set Rs=BBS.Execute("Select ID,Orders,BbsName,Url,pic,Readme,pass,Ispic From [link] where pass=1 and IsIndex=1 order by ispic,orders")
		If Rs.Eof Then
			Exit Sub
		Else
			Arr_Rs=Rs.GetRows
			Rs.Close
			Set Rs=Nothing
			ii = 0
			For i=0 To Ubound(Arr_Rs,2)
			  If Arr_rs(4,i) <> "" and Arr_rs(7,i)=0 Then
			    ii = ii + 1
			    TempText = TempText & "<a target=""_blank"" href="""&Arr_rs(3,i)&""" title="""&Arr_rs(5,i)&""">"&Arr_rs(2,i)&"</a>&nbsp;&nbsp;"
			  Else
			    TempPic = TempPic & "<a target=""_blank"" href="""&Arr_rs(3,i)&"""><img src="""&Arr_rs(4,i)&""" border=0 title="""&Arr_rs(5,i)&""" width=""88"" height=""31""></a>&nbsp;&nbsp;"
			  End If
			Next
			If i >= 9 Then TempPic = "<marquee scrollamount=3 onmouseover=stop() onmouseout=start()>"&TempPic&"</marquee>"
			If TempPic <> "" Then TempPic = "<tr><td>"&TempPic&"</td></tr>"
			If TempText <> "" Then TempText = "<tr><td>"&TempText&"</td></tr>"
			Temp="<table border=""0"" width=""94%"" cellpadding=""0"" cellspacing=""5"">"&TempText&TempPic&"</table>"
			BBS.Cache.add "linkinfo",Temp,dateadd("n",10000,now)
		End If
	End If
	S=BBS.ReadSkins("��̳����")
	S=Replace(S,"{����}",Temp)
	Response.Write S
End Sub

Function ShowNewHot()
BBS.ShowTable"����<span style='margin-left:50%'>����</span>","<table width='99%' border='0' cellspacing='0' cellpadding='3'><tr><td style='border-right:1px "&BBS.SkinsPIC(0)&" dotted' width='50%'>"&GetNewTopic(1,5)&"</td><td width='50%'>"&GetNewTopic(2,5)&"</td></tr></table>"
End Function
Function GetNewTopic(flag,Num)
Dim Rs,Sql,Noshow,i,S
Noshow=BBS.NoShowTopic()
If BBS.Cache.valid("IndexNewTopic"&Flag) then
  GetNewTopic=BBS.Cache.Value("IndexNewTopic"&Flag)
Else
  If Noshow="" Then NoShow="0"
  S="":I=0
  If Flag=1 Then
  Sql="select TopicID,Name,Face,Caption,boardid,lasttime,SqlTableID From [topic] where isdel=0 And BoardID not in("&Noshow&")  order by lasttime DESC"
  Else
  Sql="select TopicID,Name,Face,Caption,BoardID,LastTime,SqlTableID From [topic] where isdel=0 And BoardID not in("&Noshow&") And DATEDIFF('d',[LastTime],'"&BBS.NowBbsTime&"')<7 order by ReplyNum DESC"
  End If
  Set Rs=BBS.Execute(Sql)
  Do while not Rs.eof
  I=I+1
  If I>Int(Num) Then Exit Do 
  S=S&"<tr><td width=1><img src=""pic/face/"&Rs("face")&".gif""></td><td><a href=""topic.asp?boardid="&Rs("BoardID")&"&id="&Rs("topicid")&"&tb="&Rs("SqlTableID")&""">"&BBS.Fun.StrLeft(Rs("Caption"),40)&"</a></td><td><a href=userinfo.asp?name="&Rs("Name")&">"&Rs("Name")&"</a></td></tr>"
  Rs.movenext
  Loop
  Rs.Close
  GetNewTopic="<table width='99%' border='0' cellspacing='0' cellpadding='3'>"&S&"</table>"
  BBS.Cache.add "IndexNewTopic"&Flag,GetNewTopic,dateadd("n",20,now)
End If
End Function
%>