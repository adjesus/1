<!--#include file="Inc.asp"-->
<%
Dim UserName,Page_Url
If Not BBS.FoundUser Then BBS.GoToErr(10)
UserName=request.querystring("name")
If UserName <> "" Then
  Page_Url = "?Name="&UserName
Else
  Page_Url = ""
End If
If Trim(UserName)="" Then UserName=BBS.MyName
If Not BBS.Fun.CheckName(UserName) then BBS.GoToErr(1)
If Lcase(UserName)=lCase(BBS.MyName) Then
	BBS.Position=BBS.Position&" -> <a href='userinfo.asp'>�û��������<a>"
	BBS.Head "userinfo.asp"&Page_Url,"","�鿴��������" 
	Response.Write BBS.ReadSkins("�û��������")
Else
	If SESSION(CacheName& "MyGradeInfo")(21)="0" Then BBS.GoToErr(74)
	BBS.Head "userinfo.asp"&Page_Url,"","�鿴�û�����"
End If
Showuserinfo()
ShowUserBBS()
BBS.Footer()
Set BBS =Nothing

Sub Showuserinfo()
	Dim Caption,Content
	Dim Rs,Grade,UserPic,UserSex,RegIP,LastIP
	SET Rs=BBS.Execute("Select Name,Sex,Birthday,Mail,Home,IsQQpic,QQ,Pic,Pich,Picw,RegIp,LastIp,EssayNum,GoodNum,Mark,GradeID,Coin,GameCoin,BankSave,RegTime,LastTime,IsShow,IsDel,IsVip,LoginNum,Honor,Sign,Faction From[user]where Name='"&UserName&"' And Isdel=0")
	If Rs.eof then BBS.GoToErr(79)
	If Rs("IsQQpic")="1" Then
		UserPic="<img src='http://qqshow-user.tencent.com/"&Rs("QQ")&"/10/'>"
	Else
		UserPic="<img src="&BBS.Fun.GetSqlStr(Rs("pic"))&" width="&Rs("picw")&" height="&Rs("pich")&" >"
	End If
	IF Rs("Sex")=1 Then UserSex="��" Else UserSex="Ů"
	
	Grade=BBS.GetGradeInfo(Rs("GradeID"))
	Grade=split(Grade,"|")

	IF SESSION(CacheName& "MyGradeInfo")(42)="1" Then
		RegIP=BBS.Fun.GetSqlStr(Rs("RegIp"))
		LastIP=BBS.Fun.GetSqlStr(Rs("LastIp"))
	Else 
		RegIP="����":LastIP="����"
	End If
	Caption="�û���Ϣ"
	Content="<div>"&_
	"<div style=""float:left;width:180px;text-align:center"">"&UserPIc&"<br /><a href='sms.asp?action=write&name="&UserName&"'><img src='Images/Icon/sms.gif' border='0' /> ��������</a></div>"&_
	"<div style=""margin-left:180px"">"&_
	"<div style='float:left;width:40%;'><div style='margin:5px 0px;border:1px solid "&BBS.SkinsPIC(0)&";'><div class=""title1"">������Ϣ</div>"&_
	"<ul><li>�ǳƣ�"&UserName&"</li><li>�Ա�"&UserSex&"</li><li>���գ�"&Rs("Birthday")&"</li><li>���䣺<script>mail('"&BBS.Fun.GetSqlStr(Rs("Mail"))&"')</script></li><li> QQ��"&BBS.Fun.GetSqlStr(Rs("QQ"))&"</li><li>��վ��<a href="&BBS.Fun.GetSqlStr(Rs("Home"))&">�ҵ���ַ</a></li><li>ע��ʱ�䣺"&Rs("RegTime")&"</li><li>�����ʣ�"&Rs("LastTime")&"</li><li>ע��ʱIP��"&RegIP&"</li><li>������IP:"&LastIP&"</li></ul></div></div>"&_
	"<div style=""float:right;width:40%;""><div style='margin:5px 8px 5px 0px; border:1px solid "&BBS.SkinsPIC(0)&";'><div class=""title1"">��̳��Ϣ</div>"&_
	"<ul><li>��̳�ȼ���"&Grade(2)&"</li><li>ͷ�ֳƺţ�"&BBS.Fun.GetSqlStr(Rs("Honor"))&"</li><li>��̳���ɣ�"&BBS.Fun.GetSqlStr(Rs("Faction"))&"</li><li>����������"&Rs("EssayNum")&"</li><li>����������"&Rs("GoodNum")&"</li><li>"&BBS.Info(120)&"��"&Rs("Coin")&"</li><li>���д�"&Rs("BankSave")&"</li><li>"&BBS.Info(122)&"��"&Rs("GameCoin")&"</li><li>"&BBS.Info(121)&"��"&Rs("Mark")&"</li><li>��½������"&Rs("LoginNum")&"��</li></ul></div></div></div></div>"
	Rs.Close
	BBS.ShowTable Caption,Content
End Sub 

Sub ShowUserBBS()
	Dim Rs,I,SysInfo,ReplyInfo,NoShow
	NoShow=BBS.NoShowTopic()
	Set Rs=BBS.Execute("select Top 5 Face,BoardID,Caption,LastTime,TopicID,Name,SqlTableID from [Topic] where Name<>'"&UserName&"' And IsDel=0 and TopicID in (Select ReplyTopicID from [Bbs"&BBS.TB&"] where name='"&UserName&"' And IsDel=0) order by LastTime desc")
	 Do While not Rs.Eof
		If InStr(","&NoShow&",",","&Rs("BoardID")&",")=0 or (lcase(UserName)=lcase(BBS.MyName) or BBS.MyAdmin=9) Then
			ReplyInfo=ReplyInfo& "<div style=""padding:5px;border-bottom:1px dashed #ccc""><a target='_blank' href='topic.asp?boardid="&Rs("BoardID")&"&ID="&Rs("TopicID")&"&TB="&Rs("SqlTableID")&"'><img src='pic/face/"&Rs("Face")&".gif' align='absmiddle' border='0'> "&BBS.Fun.StrLeft(Rs("Caption"),50)&"</a></div>"
		End If
		Rs.MoveNext
	 Loop
	Rs.Close
	ReplyInfo="<div style='text-align:left;margin-left:450;'><div style=""padding:5px;border-bottom:1px dashed #ccc;BACKGROUND:"&BBS.SkinsPIC(2)&"""><img src='Images/icon/inn.gif' align='absmiddle'> <b>������������</b></div>"&ReplyInfo&"</div>"
	Set Rs=BBS.Execute("select Top 5 Face,BoardID,Caption,AddTime,TopicID,Name,SqlTableID from  [Topic] where name='"&UserName&"' And IsDel=0 order by AddTime desc")
	 Do While not Rs.Eof
		If InStr(","&NoShow&",",","&Rs("BoardID")&",")=0 or (lcase(UserName)=lcase(BBS.MyName) or BBS.MyAdmin=9) Then
			SysInfo=SysInfo& "<div style=""padding:5px;border-bottom:1px dashed #ccc""><a target='_blank' href='topic.asp?boardid="&Rs("BoardID")&"&id="&Rs("TopicID")&"&tb="&Rs("SqlTableID")&"'><img src='pic/face/"&Rs("Face")&".gif' align='absmiddle' border='0'> "&BBS.Fun.StrLeft(Rs("Caption"),50)&"</a></div>"
		End If
		Rs.MoveNext
	 Loop
	Rs.Close
	SysInfo="<div style='float:left;text-align:left;width:450'><div style=""padding:5px;border-bottom:1px dashed #ccc;BACKGROUND: "&BBS.SkinsPIC(2)&"""><img src='Images/icon/inn.gif' align='absmiddle'> <b>������������</b></div>"&SysInfo&"</div>"
	BBS.ShowTable UserName&" ������Ϣ",SysInfo&ReplyInfo&"<div style='clear:both'></div>"
End Sub
%>	