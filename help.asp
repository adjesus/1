<!--#include file="inc.asp"-->
<%
Dim Content,Action,Page_Url
Action=Request.querystring("Action")
If Action <> "" Then
  Page_Url = "?action="&Action
Else
  Page_Url = ""
End If
If Action="mygrade" then
	BBS.Position=BBS.Position&" -> <a href='userinfo.asp'>�û��������<a>"
	BBS.Head "help.asp"&Page_Url,"","�鿴�ȼ�Ȩ��"
Else
	BBS.Head "help.asp"&Page_Url,"","�鿴��̳����"
End if
If Len(Action)>13 Then BBS.GotoErr(1)

Select Case lcase(Action)
Case "what"
	what
Case "upload"
	Upload
Case "sms"
	Sms
Case"forget"
	Forget
Case"grade"
	Grade	
Case"usersetup"
	UserSetup
Case"say"
	Say
Case"ubb"
	Ubb
Case"BBS"
	Info
Case"gradestring","mygrade"
GradeString
Case Else
	Main
End Select
BBS.Footer()
Set BBS =Nothing

Sub Main
Content="<div align='center'><b>====== ��̳����Ŀ¼ ======</b><table border='0' cellspacing='10' cellpadding='0'><tr><td><a href=?action=what>��Լ��̳�ĳ���˵��</a></td><td><a href=?action=grade>�����û��ȼ��İ���</a></td><td><a href=?action=forget>������������İ���</a></td><td><a href=?action=upload>�����ļ��ϴ��İ���</a></td></tr><tr><td><a href=?action=sms>������̳����İ���</a></td><td><a href=?action=usersetup>���ڸ�����Ϣ�İ���</a></td><td><a href=?action=Say>���ڷ������ӵİ���</a></td><td><a href=?action=ubb>����UBB���ܵİ���</a></td></tr></table></div>"
BBS.ShowTable "��̳����",Content
End Sub



Sub GradeString()
Dim Rs,ID,GradeName,Pic,Spic,Grouping,EssayNum,Title
Dim S,GS,T,Y,N
ID=BBS.Checknum(request.querystring("ID"))
Title="�ȼ�Ȩ��"
If ID=0 Then
	If Not BBS.FoundUser Then
		BBS.GoToerr(26)
	Else
		Title="�ҵĵȼ�Ȩ��"
		Response.Write BBS.ReadSkins("�û��������")
		ID=SESSION(CacheName & "MyInfo")(15)
	End If
End If
Y="<span style='color:#F00'>��</span>"
Set Rs=BBS.Execute("Select ID,GradeName,PIC,Spic,Strings,EssayNum,Grouping,Flag From [Grade] where ID="&ID&"")
If Not Rs.eof Then
	Gs=Split(Rs(4),"|")
	GradeName=Rs(1)
	Pic=BBS.Fun.GetSqlStr(Rs(2))
	Spic=BBS.Fun.GetSqlStr(Rs(3))
	Grouping=Rs(6)
	EssayNum=BBS.Fun.GetSqlStr(Rs(5))
Else
	BBS.GotoErr(1)
End If
Rs.close
Set Rs=Nothing
S=BBS.Row("�ȼ�����:","<b>"&GradeName&"</b>","65%","")
If Grouping=0 Then S=S&BBS.Row("����ﵽ������",EssayNum&" ƪ","65%","")
If len(Pic)>3 Then T="<img src='Pic/Grade/"&pic&"' />" Else T="��"
S=S&BBS.Row("�ȼ�ͼƬ��",T,"65%","18px")
If len(SPic)>3 Then T="<img src='Pic/Grade/"&spic&"' />" Else T="��"
S=S&BBS.Row("��ݱ�־ͼƬ",T,"65%","18px")
S=S&"<div class='title'>����Ȩ��</div>"
S=S&BBS.Row("������ʾ������ɫ","<div style=""width:20px;height:20px;BACKGROUND:"&Gs(0)&""">&nbsp;</div>","65%","")
If Gs(1)="1" Then T=Y Else T="��"
S=S&BBS.Row("�Ƿ�����޸��Լ����ϣ�",T,"65%","")
If Gs(2)="1" Then T=Y Else T="��"
S=S&BBS.Row("�Ƿ�����Զ���ͷ�Σ�",T,"65%","")
S=S&BBS.Row("�����������ַ�����","<font color='#FF0000'>"&Gs(3)&"</font> �ֽ�","65%","")
If Gs(4)="1" Then T=Y Else T="��"
S=S&BBS.Row("�Ƿ���Է�����Ŀ���⣺",T,"65%","")
If Gs(5)="1" Then T=Y Else T="��"
S=S&BBS.Row("�Ƿ���Բμ�ͶƱ���",T,"65%","")
If Gs(6)="1" Then T=Y Else T="��"
S=S&BBS.Row("�Ƿ���Է���ͶƱ���⣺",T,"65%","")
If Gs(8)="1" Then T=Y Else T="��"
S=S&BBS.Row("�Ƿ�����ϴ���",T,"65%","")
S=S&BBS.Row("һ����ϴ�������","<font color='#FF0000'>"&GS(9)&"</font> ��","65%","")
S=S&BBS.Row("ÿ���ϴ���С��","<font color='#FF0000'>"&GS(10)&"</font> KB","65%","")
If Gs(11)="1" Then T=Y Else T="��"
S=S&BBS.Row("�Ƿ�����ϴ�ͷ��",T,"65%","")
S=S&BBS.Row("��̳�������������","<font color='#FF0000'>"&GS(12)&"</font> ��","65%","")
S=S&BBS.Row("����ÿ�췢���ż��Ĵ�����","<font color='#FF0000'>"&GS(7)&"</font> ��","65%","")
S=S&BBS.Row("����ÿ�����ַ���","<font color='#FF0000'>"&GS(13)&"</font> �ֽ�","65%","")
If Gs(14)="1" Then T=Y Else T="��"
S=S&BBS.Row("�Ƿ����������̳��",T,"65%","")
If Gs(15)="1" Then T=Y Else T="��"
S=S&BBS.Row("�Ƿ���Բ鿴������Ϣ��",T,"65%","")
If Gs(16)="1" Then T=Y Else T="��"
S=S&BBS.Row("���Բ���ʱ�����Ʊ༭�Լ����ӣ�",T,"65%","")
If Gs(17)="1" Then T=Y Else T="��"
S=S&BBS.Row("��������ɾ���Լ������ӣ�",T,"65%","")
S=S&"<div class='title'>����Ȩ��</div>"
If Gs(18)="1" Then T=Y Else T="��"
S=S&BBS.Row("���Ա༭���ӣ�",T,"65%","")
If Gs(19)="1" Then T=Y Else T="��"
S=S&BBS.Row("�༭���ӿ��Բ������뼣��",T,"65%","")
If Gs(20)="1" Then T=Y Else T="��"
S=S&BBS.Row("����ɾ�����ӣ�",T,"65%","")
If Gs(21)="1" Then T=Y Else T="��"
S=S&BBS.Row("�����������ӣ�",T,"65%","")
If Gs(22)="1" Then T=Y Else T="��"
S=S&BBS.Row("�����ƶ����ӣ�",T,"65%","")
If Gs(23)="1" Then T=Y Else T="��"
S=S&BBS.Row("�����������⣺",T,"65%","")
If Gs(24)="1" Then T=Y Else T="��"
S=S&BBS.Row("���Գ������⣺",T,"65%","")
If Gs(25)="1" Then T=Y Else T="��"
S=S&BBS.Row("����(��/��)�ö����⣺",T,"65%","")
If Gs(26)="1" Then T=Y Else T="��"
S=S&BBS.Row("����(��/��)���ö����⣺",T,"65%","")
If Gs(27)="1" Then T=Y Else T="��"
S=S&BBS.Row("����(��/��)���ö����⣺",T,"65%","")
If Gs(28)="1" Then T=Y Else T="��"
S=S&BBS.Row("����(��/��)�������⣺",T,"65%","")
If Gs(29)="1" Then T=Y Else T="��"
S=S&BBS.Row("����(��/��)�������⣺",T,"65%","")
If Gs(30)="1" Then T=Y Else T="��"
S=S&BBS.Row("���Խ�����������������",T,"65%","")
If Gs(31)="1" Then T=Y Else T="��"
S=S&BBS.Row("���Բ���ҪͶƱ�ɲ�ͶƱ���飺",T,"65%","")
If Gs(32)="1" Then T=Y Else T="��"
S=S&BBS.Row("���Ա༭ͶƱ��ѡ�",T,"65%","")
If Gs(33)="1" Then T=Y Else T="��"
S=S&BBS.Row("���Բ������������ƣ�",T,"65%","")
If Gs(34)="1" Then T=Y Else T="��"
S=S&BBS.Row("���Է�����̳���棺",T,"65%","")
If Gs(35)="1" Then T=Y Else T="��"
S=S&BBS.Row("����ɾ��������¼��",T,"65%","")
S=S&"<div class='title'>�߼�����Ȩ��</div>"
If Gs(36)="1" Then T=Y Else T="��"
S=S&BBS.Row("���Բ鿴�û�IP��",T,"65%","")
If Gs(37)="1" Then T=Y Else T="��"
S=S&BBS.Row("���Բ鿴��̳��־��",T,"65%","")
'If Gs(38)="1" Then T=Y Else T="��"
'S=S&BBS.Row("�������������������⣺",T,"65%","")
If lcase(Action)<>"mygrade" Then
	S=S&"<div style="" padding:5px;BACKGROUND: "&BBS.SkinsPIC(2)&";"" align=""center""><a href='javascript:history.go(-1)'>�����ء�</a><a href=help.asp>�����ذ���Ŀ¼��</a></div></form>"
End If
BBS.ShowTable Title,S
End Sub


Sub Grade()
	Dim ARs,i,S,AdminGrade,EssayGrade,otherGrade,EssayNum,Temp
	ARs=BBS.SetGradeInfoCache()
	For i=0 To Ubound(ARs,2)
	If ARs(1,i)=0 Then EssayNum="����ﵽ������<span style='color:#F00'>"&ARs(5,i)&"</span>" Else EssayNum=""
	Temp=ARs(4,i)
		If Temp<>"" Then Temp="<img src='Pic/Grade/"&ARs(4,i)&"' alt='' />"
        S="<div style=""text-align:center;padding:3px;""><div style=""float:left; width:20%""><a href=""?action=GradeString&ID="&Ars(0,i)&""">"&ARs(2,i)&"</a></div><div style=""float:left; width:10%"">"&Temp&"</div><div style=""float:left; width:20%""><img src='Pic/Grade/"&ARs(3,i)&"' alt='' /></div><div style=""float:left;"">"&EssayNum&"</div><div style=""clear: both;""></div></div>"
		If ARs(7,i)=2 Then
			AdminGrade=AdminGrade&S
		ElseIf ARs(7,i)=1 Then
			otherGrade=otherGrade&S
		Else
			EssayGrade=EssayGrade&S
		End If
	Next
		S="<div class='title'>ϵͳ�ȼ���</div>"&AdminGrade&"<div class='title'>����ȼ���</div>"&otherGrade&"<div class='title'>����ȼ���</div>"&EssayGrade
		S=S&"<div style=""padding:3px;BACKGROUND: "&BBS.SkinsPIC(2)&";"" align=""center""><a href='help.asp'>�����ذ���Ŀ¼��</a></div>"
	BBS.ShowTable"��̳�ȼ�",S
End sub

Sub Upload()
	Content="<br><br><div align='center'><b>====== ��̳�ϴ����� ======</b></div><blockquote><li>��������ʱ��������ܰ�Ŧ�ϵġ�������ҵ�Ҫ�ϴ����ļ���ѡ����Ȼ�������ϴ�����Ŧ�����ϴ���<li>ÿ�յ��ϴ���С�ʹ������ƣ�����̳����ÿ����Ա�ȼ�����ͬ�Ĺ涨��<li>����̳�����ϴ����ļ���"&Replace(BBS.Info(34)&"|"&BBS.Info(35),"|","��")&"<div align='center'><a href=help.asp>�����ذ���Ŀ¼��</a></div></blockquote>"
	BBS.ShowTable"��̳����",Content
End Sub

Sub Sms()
	Content="<br><br><div align='center'><b>====== ��̳������� ======</b></div><blockquote>��̳���书��Ҳ�൱�����ԣ�����ͬ�������������û���¼����������¼�ʱ�շ���Ϣ�������ࡣ<ul><li><b>������Ϣ</b>����¼�û����ܷ�����Ϣ��һ�����Լ���д�ռ������ƣ����û���������̳ע���û����������ı�������ݣ�����֧��ubb��ʽ������һ��������̳�в鿴���ӵ�ʱ��ֱ�Ӹ����߷�����Ϣ����Ҫ��д����ͬ��㡰���͡�����<li><b>�ռ���</b>����¼��̳�����Ϸ����û����֡��µġ����԰塱���г������Ѷ���δ������Ϣ������⡢�����ˣ����Խ��ж�ȡ��Ϣ��ȫ��ɾ��������<li><b>����Ϣ</b>����¼��̳��ÿ���б��˸��㷢���µĶ���Ϣ������ԭ������Ϣ��δ��ȡ����̳��������ʾ��ֱ�ӵ�����Ķ���</ul><div align='center'><a href=help.asp>�����ذ���Ŀ¼��</a></div></blockquote>"
	BBS.ShowTable"��̳����",Content
End sub

Sub Forget()
	Content="<br><br><div align='center'><b>====== ������������İ��� ======</b></div><blockquote><ul><li>����������룬�������ע��ʱ��д������������������<a href=UserSetup.asp?action=ForgetPassword>ȡ������</a>��</ul><ul><li>�����������������ʾ������𰸣�������̳����Ա��ϵ��������������Ϊ���趨�µ����롣</ul><div align='center'><a href=help.asp>�����ذ���Ŀ¼��</a></div></blockquote>"
	BBS.ShowTable"��̳����",Content
End sub

Sub What()
Content="<br><br><div align='center'><b>====== ������� ======</b></div>"&_
"<ul><li><b>�������ܼ�����̳��</b></li>"&_
"<ul><li>�����Ե����̳�����ġ���Աע�ᡱע��Ϊ��վ��Ա���뽫��������������д��E-mai��ַ������ȷ��Ч���Ա���������ע���ʹ�������һع��ܡ�����Ȼ���Բ���ע��Ϊ���ǵĻ�Ա������Ϊ�����ܹ�ʹ�ñ���̳��ȫ�����ܣ������Խ�����ע�ᡣ</li></ul></ul>"&_
"<ul><li><b>"&BBS.Info(121)&"��"&BBS.Info(122)&"��ʲô��? �ҵĽ�Ǯ�ߺͻ��ָ���ʲô�ô���</b></li>"&_
"<ul><li>�����ý�Ǯֵ�ͻ��ֵ�������Ծ��̳�����ա�</li>"&_
"<li>���ִ�������̳��ݵ����һ��ӵ�н϶�Ļ��ֵĻ�Ա�������̳�Ĺ��׳̶Ƚ϶ࡣ</li>"&_
"<li>��Ǯ��������̳��������ң�ӵ�н϶�Ľ�Ǯ�Ļ�Ա��������̳�����һЩ���ֲ����</li>"&_
"<li>�е���������Ҫ��һ����Ǯ�ͻ��ֲſ�������ģ���Ҳ�Ƕ�ĳЩ���������ѵ����أ����һ������¡�</li>"&_
"<li>��Ǯֵֻ��˵�������ڱ���̳�Ļ�Ծ���������һ�����������κη���ĸ���ˮƽ��</li>"&_
"<li>���ǲ���������ѻ�������Ĳ�ͬ���������ѱ������κη�ʽ���Ŵ������ӡ�</li>"&_
"<li>��������ȷ�Դ���Ǯ��������������ʽ����Ҫ�����ˮ��ƭȡ��Ǯ������������ˮ��������Ϊ�������Ķ��⹥����</li>"&_
"<li>���ǽ��Զ����ˮ�����ѽ��д�����ɾ���ʺţ���������һ����ȡ�ж���Ȩ����</li></li></ul></ul>"&_
"<ul><li><b>��ֵ����μ���ģ�</b></li>"&_
"<ul>"&_
"<li>����һƪ����"&BBS.Info(120)&"��"&BBS.Info(102)&"��"&BBS.Info(121)&"��"&BBS.Info(103)&"��"&BBS.Info(122)&"��"&BBS.Info(104)&"���ظ��������ӽ�Ǯ30(ͬʱ���Ӹ�������������ͬ�Ľ�Ǯ30)</li>"&_
"<li>���ӱ�����Ա���߰�����Ϊ�����󣬷����˷�����"&BBS.Info(120)&"��"&BBS.Info(99)&"��"&BBS.Info(121)&"��"&BBS.Info(100)&"��"&BBS.Info(122)&"��"&BBS.Info(101)&"��ȡ������������Ӧ���٣�</li>"&_
"<li>���ӱ�����Ա���߰�����Ϊ�ö��󣬷����˷�����"&BBS.Info(120)&"��"&BBS.Info(96)&"��"&BBS.Info(121)&"��"&BBS.Info(97)&"��"&BBS.Info(122)&"��"&BBS.Info(98)&"��ȡ���ö�������Ӧ���٣�</li>"&_
"<li>���ӱ�����Ա��Ϊ���ö��󣬷�����"&BBS.Info(120)&"��"&BBS.Info(93)&"��"&BBS.Info(121)&"��"&BBS.Info(94)&"��"&BBS.Info(122)&"��"&BBS.Info(95)&"��ȡ�����ö�������Ӧ���٣�</li>"&_
"<li>���ӱ�����Ա���ö��󣬷�����"&BBS.Info(120)&"��"&BBS.Info(90)&"��"&BBS.Info(121)&"��"&BBS.Info(91)&"��"&BBS.Info(122)&"��"&BBS.Info(92)&"��ȡ�����ö�������Ӧ���٣�</li></ul></ul>"&_
"<ul><li><b>��������֪���Լ��Ļ��ֺͻ��ֵ����������</b></li>"&_
"<ul><li>ֻҪ����̳���ҵ��Լ����û�����������ɲ鿴�Լ��Ļ��֡�</li>"&_
"<li>������̳������������ͨ������̳�˵���<a href=userboard.asp>�û��б�</a>�鿴��</li></ul></ul>"&_
"<ul><li><b>�������������Ӻ����ǩ���� </b></li>"&_
"<ul><li>������ͨ����̳�������û����֡��¡��޸����ϡ�����ǩ��һ����д�����ĸ���ǩ����</li>"&_
"<li>ǩ��֧��UBB������ʹ��ͼƬ����ʽ��[img]ͼƬ��ַ[/img]��</li></ul></ul>"&_
"<ul><li><b>��ο����ҵ���Ҫ�����£�</b></li>"&_
"<ul><li>������ʹ����̳���������ܣ�����������̳�������ڡ���̳�˵����µġ�<a href=Search.asp>��̳����</a>����д��������������������������ڸ������·�Ҳ���԰�����������������������ѡ��ʹ�á���������������һ���������� ���ϴε��ú���������������������������������ظ������Ƚ��м���������</li></ul></ul>"&_
"<ul><li><b>ʲô�Ǿ�������������˭�������뾫�����ģ���β鿴��</b>"&_
"<ul><li>����������̳����������м�ֵ�����������ϸ߻����ݱȽ�����������ӵģ�����ͨ�������ھ��������ҵ��ܶ����õĶ���������̳��ÿ�����涼���Լ��ľ��������������ɰ��������������Խ������ϵ����Ӽ��뵽�����������ɱ༭�����ټӹ���</li>"&_
"<li>�����������Ӽ�ʹ�ڿռ����������������Ҳ���ᱻɾ�����ᱻ���ñ�����</li>"&_
"<li>ֻҪ������ذ��棬�ڰ�������Ϸ��Ϳ��Կ��������澫����������������ɲ鿴���ڡ���̳�˵�����Ҳ��һ������������������ȫ����̳�ľ������ڣ����������Դ��ַ�ʽ�鿴�������������ؾ����뵽��ذ�����ҡ�</li></ul></ul>"&_
"<ul><li><b>��Ϊ������������ʲô������ЩȨ����</b></li>"&_
"<ul><li>������Ȼ��������֮��������Ϊ�����ģ�Ը��Ϊ�����޳�����</li>"&_
"<li>ӵ��һ��Ⱥ�ڻ������ܻ�ʱ��ά����̳�ġ�</li>"&_
"<li>�������Բ�ѯ��������ӣ���������ɾ����༭��������ӡ�</li>"&_
"<li>�������԰����Ӽ��������������ö�������ö����������ͳ������ӣ����뾫������������������������湫�档</li>"&_ 
"<li>�������������������Ϸ�����ɾ�����档</li>"&_
"<li>�����̳�����˰����̳й��ܣ��ϼ����������Թ����¼���̳�����ӡ�</li></ul></ul><div align='center'><a href=help.asp>�����ذ���Ŀ¼��</a></div></blockquote>"
BBS.ShowTable"��̳����",Content
End sub

Sub UserSetup()
Content="<br><br><div align='center'><b>====== ���ĸ�����Ϣ���� ======</b></div><blockquote>"&_
"ֻ�е�¼�û����ܽ��д��������ԭע���û��������޸ģ�������̳���������֡����ҵ����޸����ϡ����룬���Ը��µ���Ϣ���£�"&_
"<li>���룺��¼��̳����"&_
"<li>Email��ַ������������ȷ�Ϸ��������ַ"&_
"<li>���գ������������������̳����������յ���ʾ���������л�Ա��"&_
"<li>������ҳ����ѡ��еĻ��������ϣ��ô�Ҽ�ʶ�£�"&_
"<li>OICQ����ѡ�Ϊ������ϵ���������ϣ�"&_
"<li>����ͷ����ѡ��ͷ���������г��֣�����������ͼƬurl��QQ������Ϊͷ��"&_
"<li>����ǩ����֧��ubb��������룬�����������µĽ�β<br><br><div align='center'><a href=help.asp>�����ذ���Ŀ¼��</a></div></blockquote>"
BBS.ShowTable"��̳����",Content
End sub

Sub Say()
Content="<br><br><div align='center'><b>====== �������Ӱ��� ======</b></div><blockquote>"&_
"<ul><li>ֻ��ע�Ტ���ѵ�½���û����ſ��Է���һ�������⣬���ǻظ��������⡣"&_
"<li>�����������İ���˵��:</B>"&_
"<ul><li>�����ִ�С��ѡ����Ҫ���ֺţ��ڳ��ֵ�����������������ݼ��ɡ�"&_
"<li>��������ɫ��ѡ����Ҫ����ɫ���ڳ��ֵ�����������������ݼ��ɡ�"&_
"<li>���ظ��ɼ�������������ֻ�лظ��˸�������û��ſɿ�����<br>�ڷ���ʱ�㡰���������ġ��ظ��ɼ�������ʱ�������ӿ��ڳ���<font color=blue>[reply]����[/reply]</font>�������еġ����ݡ��滻Ϊ���Լ������ݼ��ɡ�"&_
"<li>��ָ�����ߡ�������ֻ�б�ָ����ע���û��ſɼ���<br>�ڷ���ʱ�㡰����������ָ�����ߡ�����ʱ�������ӿ��ڳ���<font color=blue>[UserName=admin]����[/UserName]</font>��ǩ�������еġ�admin���滻Ϊĳ��ע���û����������ݡ��滻Ϊ������ݼ��ɡ�"&_
"<li>����Ǯ�ɼ�������������ֻ�дﵽָ����Ǯֵ���û��ſɿ���<br>�������Ǯ�ɼ�����ʱ�������ӿ��ڳ���<font color=blue>[COIN=1000]����[/COIN]</font>�����е�1000�滻Ϊ����Ҫ�Ľ�Ǯ�����������ݡ��滻Ϊ����Ҫ�����ݼ��ɡ�"&_
"<li>�����ֿɼ�������������ֻ�дﵽָ������ֵ���û��ſɿ���<br>�����������ֿɼ�����ʱ�������ӿ��ڳ���<font color=blue>[MAKE=3]����[/MAKE]</font>�����е�3�滻Ϊ����Ҫ�Ľ�Ǯ�����������ݡ��滻Ϊ����Ҫ�����ݼ��ɡ�"&_
"<li>�����ѿɼ����ù��ܿ������á�ʵ�ü�ֵ�ߡ������ӵ��Ķ��۸��Ķ����蹺����ܿ���<br>�ڷ���ʱ�㡰���������ġ����ѿɼ�������ʱ������һ�����������������ӵļ۸���ʱ�������ӿ��ڳ���<font color=blue>[BUYPOST=100]����[/BUYPOST]</font>��ǩ���������ݡ��滻Ϊ���Լ������ݼ��ɡ�"&_
"<li>�����ڿɼ�������������ֻ�е��˹涨���ں�ſ��Կ�����<br>�ڷ���ʱ�㡰���������ġ����ڿɼ�������ʱ������һ�����������������ӵĿɼ����ڣ���ʱ�������ӿ��ڳ���<font color=blue>[DATE=2003-10-1]����[/DATE]</font>��ǩ���������ݡ��滻Ϊ���Լ������ݼ��ɡ�"&_
"<li>���Ա�ɼ�������������ֻ��ָ�����Ա�ſ��Կ�����<br>�ڷ���ʱ�㡰���������ġ��Ա�ɼ�������ʱ������һ��������������1��0��1�����У�0����Ů������ʱ�������ӿ��ڳ���<font color=blue>[SEX=1]����[/SEX]</font>��ǩ���������ݡ��滻Ϊ���Լ������ݼ��ɡ�"&_
"<li>����½�ɼ���������ֻ�е�½���û��ſɼ���<br>�ڷ���ʱ�㡰������������½�ɼ�������ʱ�������ӿ��ڳ���<font color=blue>[LOGIN]����[/LOGIN]</font>��ǩ���������ݡ��滻Ϊ������ݼ��ɡ�</ul>"&_
"<li>ÿ���������涼�п��ٻظ����������ֱ���������������ݻظ���"&_
"<li>�����Ҫ�õ�ĳЩ��ǩ�緢���������ϴ��ļ��ȣ������������Ϸ��ġ��ظ����ӡ����ɣ��ظ�����ʱ�뷢����ͬ."&_
"<li>���˻ظ�����ʱ���������������⣨�����Ը������������ӵ�����ѡ����Ӧ�ġ������־������Ҫȷ����д�������Ա�ɹ��������ӡ�"&_
"<li>��������ܰ�Ŧ��һ���İ�Ŧ�����������ٲ���ĳЩUBB��ǩ����<a href=?action=Ubb>�������UBB����</a>��</ul></ul><div align='center'><a href=help.asp>�����ذ���Ŀ¼��</a></div>"
BBS.ShowTable"��̳����",Content
End Sub

Sub Ubb()
Content="<br><br><div align='center'><b>====== UBB��ǩ���� ======</b></div><blockquote>"&_
"<ul>UBB��ǩ���ǲ�����ʹ��HTML�﷨������£�ͨ����̳������ת��������������֧���������õġ���Σ���Ե�HTMLЧ����ʾ������Ϊ����ʹ��˵����"&_
"<p><font color=red>[B]</font><b>����</b><font color=red>[/B]</font><br>�����ֵ�λ�ÿ��������������Ҫ���ַ�����ʾΪ����Ч����"&_
"<p><font color=red>[I]</font><i>����</i><font color=red>[/I]</font><br>�����ֵ�λ�ÿ��������������Ҫ���ַ�����ʾΪб��Ч����"&_
"<p><font color=red>[U]</font><u>����</u><font color=red>[/U]</font><br>�����ֵ�λ�ÿ��������������Ҫ���ַ�����ʾΪ�»���Ч����"&_
"<p><font color=red>[align=center]</font>����<font color=red>[/align]</font><br>�����ֵ�λ�ÿ��������������Ҫ���ַ���centerλ��center��ʾ���У�left��ʾ����right��ʾ���ҡ�"&_
"<p><A HREF='http://www.74177.com/bbs'><font color=red>http://www.74177.com/bbs</font></A><br>ֱ��������ַ����̳���Զ�ʶ��"&_
"<P><font color=red>[URL=http://www.74177.com/bbs]</font><A HREF=http://www.74177.com/bbs>��Լ��̳</A><font color=red>[/URL]</font>��<br>������Ҳ�������Ӿ����ַ�����������ӡ�"&_
"<P><font color=red>[EMAIL]</font><A HREF=""mailto:abc@abc.com"">abc@abc.com</A><font color=red>[/EMAIL]</font><br>"&_
"<font color=red>[EMAIL=MAILTO:abc@abc.com]</font><A HREF=""mailto:abc@abc.com"">����</A><font color=red>[/EMAIL]</font>��<br>�����ַ������Լ����ʼ����ӣ��������Ӿ����ַ�����������ӡ�"&_
"<P><font color=red>[img]</font>http://www.74177.com/bbs/images/pic.gif<font color=red>[/img]</font><br>�ڱ�ǩ���м����ͼƬ��ַ����ʵ�ֲ�ͼЧ����"&_
"<P><font color=red>[flash]</font>Flash���ӵ�ַ<font color=red>[/Flash]</font><br>�ڱ�ǩ���м����FlashͼƬ��ַ����ʵ�ֲ���Flash��"&_
"<P><font color=red>[Code]</font>����<font color=red>[/Code]</font><br>�ڱ�ǩ��д�����ֿ�ʵ��html�б��Ч����"&_
"<P><font color=red>[quote]</font>����<font color=red>[/quote]</font><br>�ڱ�ǩ���м�������ֿ���ʵ��HTMl����������Ч����"&_
"<P><font color=red>[list]</font>����<font color=red>[/list]</font> <font color=red>[list=a]</font>����<font color=red>[/list]</font>  <font color=red>[list=1]</font>����<font color=red>[/list]</font>��<br>����list���Ա�ǩ��ʵ��HTMLĿ¼Ч����"&_
"<P><font color=red>[fly]</font>����<font color=red>[/fly]</font><br>�ڱ�ǩ���м�������ֿ���ʵ�����ַ���Ч������������ơ�"&_
"<P><font color=red>[move]</font>����<font color=red>[/move]</font><br>�ڱ�ǩ���м�������ֿ���ʵ�������ƶ�Ч����Ϊ����Ʈ����"&_
"<P><font color=red>[light]</font>����<font color=red>[/light]</font><br>�ڱ�ǩ���м�������ֿ���ʵ������������ɫ��������Ч��"&_
"<P><font color=red>[shadow=255,red,2]</font>����<font color=red>[/shadow]</font><br>�ڱ�ǩ���м�������ֿ���ʵ��������Ӱ��Ч��shadow����������Ϊ��ȡ���ɫ�ͱ߽��С��"&_
"<P><font color=red>[color=��ɫ����]</font>����<font color=red>[/color]</font><br>����������ɫ���룬�ڱ�ǩ���м�������ֿ���ʵ��������ɫ�ı䡣"&_
"<P><font color=red>[size=����]</font>����<font color=red>[/size]</font><br>�������������С���ڱ�ǩ���м�������ֿ���ʵ�����ִ�С�ı䡣"&_
"<P><font color=red>[face=����]</font>����<font color=red>[/face]</font><br>��������Ҫ�����壬�ڱ�ǩ���м�������ֿ���ʵ����������ת����"&_
"<P><font color=red>[em1]</font><br>��̳����ͼƬ���롣���е�����1��180֮����ͼƬ���롣"&_
"<P><div align='center'><a href=help.asp>�����ذ���Ŀ¼��</a></div></blockquote>"
BBS.ShowTable"��̳����",Content
End sub

%>