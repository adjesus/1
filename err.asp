<!--#include file="inc.asp"-->
<%
Dim ErrNum,NeedLogin,GoBack,Caption,Content
NeedLogin=False
GoBack=False
BBS.Head "err.asp?errnum="&request.querystring("ErrNum"),"","������Ϣ"
Caption="������Ϣ��"
ErrNum=request.querystring("ErrNum")
If Not isnumeric(ErrNum) Then ErrNum=1
Select Case ErrNum
	Case 1
		Caption="�Ƿ�������"
		Content = "<li>����ĵ�ַ���������벻Ҫ�ֶ�ȥ���ĵ�ַ��������</li>"
	Case 2
		Caption="�Ƿ����� ��"
		Content = "<li>�����ύ������������̳�ⲿ���벻Ҫ����̳�ⲿ�ύ���ݣ�лл������</li>"
	Case 3
		NeedLogin = True
		Content ="<li>�����¼��̳����ܽ��в�������<a href='login.asp'>��½</a>��</li>"
	Case 4
		Caption="�Ƿ����� ��"
		Content ="<li>�벻Ҫ������ύ��</li>"
	Case 5
		GoBack = True
		Caption="��½ʧ��"
		Content="<li>��վΪ�˷�ֹ���Ⳣ�Ի�����½��2�ε�½�������Ϊ<Font color=red>"&BBS.Info(10)&"</Font>����</li>"
	Case 6
		GoBack=True
		NeedLogin = True
		Caption ="��½ʧ��"
		Content = "<li>�װ����û��������ڵ�½ʱ��Ҫ�������û��������롣</a></li>"
	Case 7
		GoBack=True
		Caption ="��½ʧ��"
		Content = "<li>����д��֤�롣</li>"	
	Case 8
		GoBack=True
		NeedLogin = True
		Content = "<li>����д��ȷ����֤��</li>"	
	Case 9
        NeedLogin = True
		GoBack = True
		Caption="��½ʧ��"
		Content="<li>�����û������������</li><li>����˺ű���ʱɾ����</li><li>ע�⣺��������½����5�Σ������콫�����ô��ʺŵ�½��</li>"
	Case 10
		GoBack=True
		NeedLogin = True
		Caption="����ʧ�� ��"
		Content="<li>�����ܳɹ��Ľ���ð��棡</li><li>�ð���Ϊֻ��ע���Ա���Խ��룡</li><li>�㻹û��<a href=login.asp>��½</a>��</li>"
	Case 11
		GoBack=True
		Caption="����ʧ�� ��"
		Content="<li>�����ܳɹ��Ľ���ð��棡</li><li>�ð��治���ڣ�</li><li> ����ȷ���ʱ���̳�����ڻص�<a href=Index.asp>��ҳ</a>��</li>"
	Case 12
		GoBack=True
		Caption="����ʧ�� ��"
		Content="<li>�����ܳɹ��Ľ���ð��棡</li><li>�ð����ѱ�������ֹͣ����!</li>"
	Case 13
		GoBack=True
	   	Caption="����ʧ�� ��"
		Content="<li>�㲻�ܳɹ��Ľ���ð��棡</li><li> �ð���ΪVIP��̳����������VIP�û���</li>"		
	Case 14
		GoBack=True
	   	Caption="����ʧ�� ��"
		Content="<li>�㲻�ܳɹ��Ľ���ð��棡</li><li> �ð���Ϊ��֤��̳���㻹û�еð�������֤��</li>"
	Case 15
		GoBack=True
	   	Caption="����ʧ�� ��"
		Content="<li>�㲻�ܳɹ��Ľ���ð��棡</li><li> ������Ϊ���Ƶȼ�,��Ŀǰ����̳�ȼ��ﲻ���ð����Ҫ��</li>"
	Case 16
		GoBack=True
	   	Caption="����ʧ�� ��"
		Content="<li>�㲻�ܳɹ��Ľ���ð��棡</li><li> ������Ϊ���Ʒ�����,��Ŀǰ�ķ��������ﲻ���ð����Ҫ��</li>"
	Case 17
		GoBack=True
	   	Caption="����ʧ�� ��"
		Content="<li>�㲻�ܳɹ��Ľ���ð��棡</li><li> ������Ϊ���ƻ���,��Ŀǰ�Ļ��ִﲻ���ð����Ҫ��"				
	Case 18
		GoBack=True
	   	Caption="����ʧ�� ��"
		Content="<li>�㲻�ܳɹ��Ľ���ð��棡</li><li> ������Ϊ���ƽ�Ǯ,��Ŀǰ�Ľ�Ǯ�����ﲻ���ð����Ҫ��</li>"
	Case 19
		GoBack=True
	   	Caption="����ʧ�� ��"
		Content="<li>�㲻�ܳɹ��Ľ���ð��棡</li><li> ������Ϊ������Ϸ��,��Ŀǰ����Ϸ�������ﲻ���ð����Ҫ��</li>"				 
	Case 20
        NeedLogin = True
		GoBack=True
		Caption="�û�����"
		Content="<li>�㲻�ܳɹ��Ľ����ҳ�棡</li><li>��ҳ��Ϊֻ��ע���Ա���Խ��룡</li><li>�㻹û��<a href='login.asp'>��½</a></li>��"
	Case 21
		GoBack=True
		Content="<li>��Ĳ�������</li><li>�����Ӳ�����</li><li>��������Ѿ�ɾ��</li><li>����<a href='Index.asp'>��̳��ҳ</a></li>" 
	Case 22
		GoBack = True
		Content="<li>��Ĳ�������</li><li>�������Ѿ���������</li><li><a href='Index.asp'>������̳��ҳ</a></li>"
	Case 23
		Caption="ע��ʧ�� ��"
		GoBack = True
		Content = "<li>��Ǹ����̳��ͣ���û�ע�ᣡ</li>" 
	Case 24
		Caption="ע��ʧ�� ��"
		GoBack=True
		Content="<meta http-equiv=refresh content=4;url=Index.asp><li>�Բ����㲻�ܳɹ�ע�ᣡ����</li><li>����̳Ϊ�˷�ֹ����ע��ȶ���ע�ᣬͬһ�û���Դ����ע����<b>"&BBS.Info(9)&"</b> ���ӣ�</li>"
	Case 25
        NeedLogin = True
		GoBack=True
		Caption="����ʧ��"
		Content="<li>��û������������ӵ�Ȩ��</li><li>ֻ��ע���Ա�ſ��Խ��룡</li><li> �㻹û��<a href=""login.asp""> ��½ </a>��</li>"
    Case 26
        NeedLogin = True
		GoBack = True
		Content="<li>��Ŀǰ�Ĳ��ǻ�Ա��ֻ�������������ӵ�Ȩ��</li>"
	Case 27
		GoBack = True
		Content = "<li>�Բ����㲻�ܳɹ��ط������ӣ�</li><li> �㲢û����д��������ݣ�</li>"
	Case 28
		GoBack = True
		Content = "<li>�Բ����㲻�ܳɹ��ط������ӣ�</li><li>���ӱ����ַ���������̳���ƣ�</li>" 
    Case 29
		GoBack = True
		Content = "<li>�����ַ���������̳���ƣ�</li>"
 	Case 30
		GoBack=True
		Content = "<li>�Բ����㲻�ܳɹ��ط������ӣ�</li><li>��֤�벻�ԣ�����д��ȷ����֤��</li>"	
	Case 31
	    NeedLogin = True
		GoBack=True
		Content = "<li>�㻹û�� <a href=""login.asp"">��½</a> �� <a href=""register.asp"">ע��</a> ��</li>"	
 	Case 32
		GoBack=True
		Content="<li>��Ĳ�������</li><li>�����Ӳ����� �� �Ѿ�ɾ��</li><li>����<a href=Index.asp>��̳��ҳ</a></li>"
	Case 33 
		GoBack = True
		Content ="<li>��Ĳ�������</li><li>�㲻�Ǹ��������߻�ð���İ������Բ��ܱ༭������</li>"
	Case 34
		GoBack = True
		Content = "<li>�㲻�ܱ༭���ӣ�</li><li>��Ϊ�㳬����������ͨ�û��༭�Լ����ӵ�ʱ�� (��������<font color='#F00'>"&BBS.Info(12)&"</font>������)</li>"
	Case 35
		GoBack=True
		Content = "<li>���ӷ���ʧ�ܣ�</li><li>��֤��ʧЧ������д��ȷ����֤��</li>"	
	Case 36
		GoBack = True
		Content = "<li>����û����д�����ı���ѡ��</li>"
	Case 37
		GoBack = True
		Content = "<li>�벻Ҫʹ�÷Ƿ��ַ��������ڽ�ֹע��֮�У���</li>"		
	Case 38
		GoBack = True
		Content = "<li>�û��� �� ���� ����С��14���ַ�(7������)���߲����õ����ַ���</li>"		
	Case 39
		Caption="ע��ʧ�ܣ�"
		GoBack = True
		Content = "<li>ע��ʧ�ܣ�������������ע�������ǳ��Ѿ�����һ���û�ʹ���ˣ�</li>"		
	Case 40
		GoBack = True
		Content = "<li>�Բ�����������������������������!</li>"	
	Case 41
		GoBack = True
		Content = "<li>��������������벻��ͬ��</li>"		
	Case 42
		GoBack = True
		Content = "<li>����д��ȷ����Ч��EMail��ַ��</li>"
	Case 43
		GoBack = True
		Content = "<li>���������⡱������𰸡����ַ�̫�̣��������4���ַ���</li>"
	Case 44
		GoBack = True
		Content ="<li>���������⡱������𰸡��ĺ��зǷ��ַ���</li>"					   
	Case 45 
		GoBack = True
		Content ="<li>��վ���ò�����ʹ���ⲿͷ��</li>"		
	Case 46
		GoBack =True
		Content ="<li>������ѡ��QQ������Ϊͷ������ȷ��д���QQ���룡</li>"		
	Case 47
		GoBack = True
		Content ="<li>����д��һЩ��Ŀ���ַ�����������̳�����ƣ�</li>"		
	Case 48
		GoBack = True
		Content = "<li>ͷ���Ⱥ͸߶ȱ�����������д��</li>"		
	Case 49
		GoBack = True
		Content = "<li>����д�������ѱ�ע�ᣡ</li>"	
	Case 50
		GoBack = True
		Content = "<li>����⣬��������û�����ֵ����ݣ�</li>"
	Case 51	
		GoBack = True
		Content = "<li>�����Ĳ����ڱ��淢��ģ����������ܱ༭������</li>"
	Case 52
        GoBack = True
		Content = "<li>�Բ��𣬲��ܷ������ԣ����"&BBS.Info(120)&"�ﲻ��<font color=red>"&BBS.Info(123)&"</font>����</li>"
	Case 53
        GoBack = True
		Content = "<li>�Բ��𣬱�վ�趨���Լ��1����</li>"
	Case 54
        GoBack = True
		Content = "<li>����ʧ�ܣ���̳�в����ڸ����Զ���</li>"
	Case 55
        GoBack = True
		Content = "<li>����ʧ�ܣ������Ѿ����ܷ���������</li>"		
 	Case 56
        GoBack =True
		Content = "<li>����д�ľ����벻��ȷ��</li>"
	Case 57
		GoBack = True
		Content = "<li>�һ�����ʧ�ܣ�</li><li>����������ʾ���������𰸲���ȷ��</li>"
	Case 58
		GoBack=True
		Content="<li>�����Ӳ����� ���� �Ѿ�ɾ��</li><li><a href=""Index.asp"">������̳��ҳ</a></li>"
	Case 59
		GoBack = True
		Content = "<li>�������Ѿ������ö�����</li>"		
	Case 60
        GoBack = True
		Content = "<li>�������Ѿ������ö�����</li>"	
	Case 61
        GoBack = True
		Content = "<li>�����������Ѿ�ȡ�������ö���</li>"
	Case 62
        GoBack = True
		Content = "<li>��ѡ�����ͬһ�����棬�����ƶ��ˣ�</li>"	
	Case 63
        GoBack = True
		Content = "<li>�����Ĺؼ����ַ�����С����̳���Ƶ� 2 ���ַ� </li>"
	Case 64
        GoBack = True
		Content = "<li>����̳������ÿ������ʱ����Ϊ "&BBS.Info(16)&" ��</li>"
	Case 65
        GoBack = True
		Content = "<li>�Բ�����̳������ÿ��ֻ�ܿ��Խ���"&BBS.Info(49)&"����������</li>"
	Case 66
		GoBack = True
		Content = "<li>һЩѡ��������������</li>"
	Case 67
		GoBack = True
		Content = "<li>�����ڵĵȼ��鲻�ܷ����͹����棡�鿴 <a href='help.asp?action=mygrade'>�ҵ�Ȩ��</a></li>"
	Case 68
		GoBack = True
		Content = "<li>�㲻�Ǹð���İ�����</li>"
	Case 69
		GoBack = True
		Content = "<li>�Ҳ�����Ӧ�ļ�¼�������Ѿ�ɾ���ˡ�</li>"											
	Case 70
        GoBack = True
		Content = "<li>�㲻�ܽ��в�������ȷ����ĵȼ�Ȩ�ޣ��鿴 <a href='help.asp?action=mygrade'>�ҵ�Ȩ��</a></li>"					
	Case 71
        GoBack = True
		Content = "<li>�㲻�ܽ��в������㲻�Ǹð���İ�����</li>"
	Case 72
		Content = "<li>�찡���㱻����Ա�߳�����̳��</li><li>��"&BBS.Info(8)&"�������㲻�ܵ�½��̳��</li>"
	Case 73
		GoBack = True
		Content = "<li>�㲻�ܱ༭�����㷢���Ĺ��棡</li>"
	Case 74
        GoBack = True
		Content = "<li>�����ڵĵȼ��鲻�ܲ鿴�û���Ϣ����ȷ����ĵȼ�Ȩ�ޣ��鿴 <a href='help.asp?action=mygrade'>�ҵ�Ȩ��</a></li>"
	Case 75
        GoBack = True
		Content = "<li>�����ڵĵȼ��鲻�ܽ�����̳��������ȷ����ĵȼ�Ȩ�ޣ��鿴 <a href='help.asp?action=mygrade'>�ҵ�Ȩ��</a></li>"
 	Case 76
		GoBack=True
		Content="<li>�㲻�ܽ��в������㲻�ǹ���Ա���������</li>"
 	Case 77
		GoBack=True
		Content="<li>�Բ��𣡱�վû�п�����̳��������!</li>"
	Case 78
		GoBack=True
		Caption="��½ʧ�ܣ�"
		Content="<li>��ע�����Ϣ����û��ͨ������Ա����ˡ�</li><li>�����ĵȴ�����Ա����ˣ�лл������</li>"
	Case 79
		GoBack=True
		Caption="�Ҳ����û���"
		Content="<li>���û����Ͽ����Ѿ�ɾ������δ�����ͨ����</li>"
	Case 81
		GoBack = True
		Content = "<li>ͷ��ͼƬ��·�����зǷ��ַ���</li>"
	Case 82
		GoBack = True
		Content = "<li>����̳Ϊ������̳��ֻ������������ϵĻ�Ա������</li>"
	Case Else
       Content = "��~~~Ϊʲô^_^�Ǻ�"
End Select
	If GoBack Then Content=Content&"<li><a href=javascript:history.go(-1)>������һҳ</a>"
	Content="<div style=""margin:18px;line-height:150%"">"&Content&"</div>"
	BBS.ShowTable Caption,Content
	IF Needlogin Then
		Dim Temp
		Temp=Request.ServerVariables("HTTP_REFERER")
			If instr(lcase(Temp),"login.asp")>0 or instr(lcase(Temp),"err.asp")>0 then
		Else
			Session(CacheName&"BackURL")=Temp
		End If
		Temp="<form method=""post"" style=""margin:0px"" action=""login.asp?action=login"">"
		Temp=Temp&BBS.Row("<b>�����������û�����</b>","<input name=""name"" type=""text"" class=""submit"" size=""20"" /> <a href=""register.asp"">û��ע�᣿</a>","65%","")
		Temp=Temp&BBS.Row("<b>�������������룺</b>","<input name=""Password"" type=""password"" size=""20"" /> <a href=""usersetup.asp?action=forgetpassword"">�������룿</a>","65%","")
		If BBS.Info(14)="1" Then
			Temp=Temp&BBS.Row("<b>�������ұߵ���֤�룺</b>",BBS.GetiCode,"65%","")
		Else
			Temp=Temp&"<input name=""iCode"" type=""hidden"" value=""BBS"" />"
		End If
		Temp=Temp&BBS.Row("<b>Cookie ѡ�</b>","<input type=radio  name=cookies value=""0"" checked class=checkbox />������ <input type=radio  name=cookies value=""1"" class=checkbox />����һ�� <input type=radio  name=cookies value=""30"" class=checkbox />����һ��","65%","")
		Temp=Temp&BBS.Row("<b>ѡ���½��ʽ��</b>","<input type=radio value=""1"" checked name='hidden' class=checkbox />������½ <input type='radio' value='2' name='hidden' class=checkbox />�����½","65%","")
		Temp=Temp&"<div style="" padding:5px;BACKGROUND: "&BBS.SkinsPIC(1)&";"" align=""center""><input name=""submit"" type=""submit"" value=""�� ½"" /></div></form>"
		BBS.ShowTable"�û���½",Temp
	End If
BBS.Footer()
Set BBS =Nothing
%>