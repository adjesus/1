<!--#include file="Admin_Check.asp"-->
<%
Dim Action
Head()
Action=lcase(Request.querystring("Action"))
Select case Action
Case"bbsinfo"
	CheckString "01"
	BbsInfo
Case"configdata"
	CheckString "02"
	ConfigData
Case"gapad"
	CheckString "04"
	GapAd
Case"a_e_link"
	CheckString "05"
	A_E_Link
Case"lockip"
	CheckString "06"
	LockIp
Case"addlockip","editip"
	CheckString "06"
	A_E_LockIp
Case"clean"
	CheckString "08"
	Clean
Case"topadmin"
	'CheckString "22"
	TopAdmin
Case"boardadmin"
	CheckString "23"
	BoardAdmin
Case"grade"
	CheckString "27"
	Grade
Case"a_e_grade"
	CheckString "27"
	A_E_Grade
Case"delessay"
	CheckString "31"
	delessay
Case"moveessay"
	CheckString "32"
	MoveEssay
Case"delsms"
	CheckString "34"
	delsms
Case"menu"
	CheckString "41"
	Menu
Case"addmenu"
	CheckString "41"
	AddMenu
Case"editmenu"
	CheckString "41"
	EditMenu
Case"bank"
	CheckString "44"
	Bank
Case"faction"
	CheckString "45"
	Faction
Case"a_e_faction"
	CheckString "45"
	A_E_Faction	
Case"sqltable"
	CheckString "54"
	SqlTable
Case"updatebbs"
	CheckString "55"
	UpdateBbs
Case"spacesize"
	CheckString "57"
	SpaceSize
End select
Footer()


Function GoForm(Str)
	GoForm="<form method=POST  name=form style='margin:0' action='Admin_Confirm.asp?Action="&Str&"'>"
End Function

Sub BbsInfo
	Dim S,Info
	Set Rs=BBS.execute("Select Info From [config]")
	Info=split(Rs("Info"),",")
	Rs.Close
	With Response
	.Write"<a name='inc'></a><div class='mian'><div class='top'>��̳ϵͳ����</div><div class='divth'>��<a href='#inc1'>��̳������Ϣ</a>����<a href='#inc2'>��̳��ʾ����</a>����<a href='#inc3'>�ϴ�����</a>�� ��<a href='#inc4'>�û�ѡ��</a>�� ��<a href='#inc5'>��������</a>�� ��<a href='#inc6'>��̳��Դ����</a>��</div></div>"
'1-10������Ϣ
	.Write GoForm("BbsInfo")&"<a name='inc1'></a><div class='mian'><div class='top'>��̳������Ϣ</div>"
	DIVTR"�ر���̳��","ά���ڼ�����ùر���̳",GetRadio("info3","����",Info(3),0)&GetRadio("info3","�ر�",Info(3),1),40,1
	DIVTR"�ر���̳��ʾ��Ϣ","���ùر���̳����ʾ����Ϣ,֧��Html�﷨","<textarea rows='3' name='info4' cols='70'>"&Info(4)&"</textarea>",55,2
	DIVTR"��̳����","�����̳����","<input type='text' class='text' name='info0' size='40' value='"&Info(0)&"' /> ",40,1
	DIVTR"��̳��ַ��","��̳�ķ��ʵ�ַ","<input type='text' class='text' name='info1' size='40' value='"&Info(1)&"' />β����Ҫ�ӡ�/��",40,2
	DIVTR"��ҳ��ַ��","��ҳ�ķ��ʵ�ַ,���û�пɲ���","<input type='text' class='text' name='info2' size='40' value='"&Info(2)&"' />",40,1
	DIVTR"��վ���ڣ�","��̳��ɿ�ҵ������","<input type='text' class='text' name='info5'  value='"&Info(5)&"' /> (��ʽ��YYYY-M-D)",40,2
	DIVTR"��̳������棺","֧��Html�﷨,���ģ����ʾ����Ϊ{���}","<textarea rows='3' name='info6' cols='70'>"&Info(6)&"</textarea>",58,1
	DIVTR"��̳��Ȩ��Ϣ��","��̳�ײ���Ϣ,֧��Html�﷨","<textarea rows='3' name='info7' cols='70'>"&Info(7)&"</textarea>",55,2
	DIVTR"����������ʱ��","�趨��������������ʱ��","<input type='text' class='text' name='info8' size='5' value='"&Info(8)&"' />����",40,1
	DIVTR"ע������","ͬһ��Դ��ע����ʱ��,�������ʹ�������, ������Ϊ0","<input type='text' class='text' name='info9' size='5' value='"&Info(9)&"' />����",55,2
	DIVTR"��½�����","ͬһ��Դ�ĵ�½���ʱ��,�������ʹ�������, ������Ϊ0","<input type='text' class='text' name='info10' size='5' value='"&Info(10)&"' />����",55,1
	DIVTR"���������","ͬһ��Դ�ķ������ʱ��,�������ʹ�������, ������Ϊ0","<input type='text' class='text' name='info11' size='5' value='"&Info(11)&"' />��",55,2
	DIVTR"�༭ʱ�䣺","��ͨ��Ա�޸��Լ�������Чʱ�䣬�������ʹ�������, ������Ϊ0","<input type='text' class='text' name='info12' size='5' value='"&Info(12)&"' />����",55,1
	DIVTR"�������ʱ�䣺","����ÿ��������ʱ����,����Ա���ܴ���","<input type='text' class='text' name='info17' size='5' value='"&Info(17)&"' /> ��",40,2
	DIVTR"ע����֤�룺","",GetRadio("info13","��",Info(13),0)&GetRadio("info13","��",Info(13),1),30,1
	DIVTR"��½��֤�룺","",GetRadio("info14","��",Info(14),0)&GetRadio("info14","��",Info(14),1),30,2
	DIVTR"������֤�룺","",GetRadio("info15","��",Info(15),0)&GetRadio("info15","��",Info(15),1),30,1
	DIVTR"ɾ����̳��־��","",GetRadio("info16","�ֹ�ɾ��",Info(16),0)&GetRadio("info16","�Զ�ɾ��7��ǰ�ļ�¼",Info(16),1),30,2
'20-29��ʾ����
	.Write"</div><a name='inc2'></a><div class='mian'><div class='top'>��̳��ʾ����<a href='#inc'>��</a></div>"
	DIVTR"��ʾϵͳ��Ϣ��","������ҳ���桢���ٵ�½",GetRadio("info20","��",Info(20),0)&GetRadio("info20","��",Info(20),1),40,1
	DIVTR"��ʾ���룺","��ѡ���-�������",GetRadio("info21","��",Info(21),0),40,2
	DIVTR"��ʾ��Ա���գ�","�Ƿ���ʾ��ҳ�Ļ�Ա������Ϣ",GetRadio("info22","��",Info(22),0)&GetRadio("info22","��",Info(22),1),40,1
	DIVTR"��ʾ��̳���ˣ�","�Ƿ���ʾ��̳��ҳ����������",GetRadio("info23","��",Info(23),0)&GetRadio("info23","��",Info(23),1),40,2
	DIVTR"��ʾ���ݲ�ѯ��","�Ƿ���ʾ��̳�ײ������ݲ�ѯ",GetRadio("info24","��",Info(24),0)&GetRadio("info24","��",Info(24),1),40,1
	DIVTR"��ʾִ��ʱ�䣺","�Ƿ���ʾҳ���²��ļ���ʱ��",GetRadio("info25","����ʾ",Info(25),0)&GetRadio("info25","�Ժ�����ʾ",Info(25),1)&GetRadio("info25","������ʾ",Info(25),2),40,2
	DIVTR"��ʾ��������","������̳�ķ��ʼ�����",GetRadio("info26","ʧЧ",Info(26),0)&GetRadio("info26","��ʾ",Info(26),1)&GetRadio("info26","����ʾ",Info(26),2),40,1
	DIVTR"��ʾλ�õ�����","�Ƿ���ʾ[���λ��]������<br />(������������˵�)","<br>"&GetRadio("info27","��",Info(27),0)&GetRadio("info27","��",Info(27),1),60,2
'30-39�ϴ�����
	.Write"</div><a name='inc3'></a><div class='mian'><div class='top'>�ϴ�����<a href='#inc'>��</a></div>"
	DIVTR"�ļ��ϴ���","�Ƿ��������ϴ�",GetRadio("info30","��ֹ",Info(30),0)&GetRadio("info30","����",Info(30),1),40,1
	DIVTR"��������","�Ƿ����ϴ��ļ�������",GetRadio("info31","��",Info(31),0)&GetRadio("info31","��",Info(31),1),40,2
	DIVTR"�ļ����ؼ�����","�Ƿ����ϴ��ļ������ؼ�����",GetRadio("info32","��",Info(32),0)&GetRadio("info32","��",Info(32),1),40,1
	DIVTR"�������������ػ���ʾ��","�Ƿ�������������������ʽ��",GetRadio("info38","��",Info(38),0)&GetRadio("info38","��",Info(38),1),40,1
	DIVTR"ͷ���ϴ���С��","����ͷ������ϴ��Ĵ�С","<input type='text' class='text' name='info33' size='5' value='"&Info(33)&"' /> KB",40,2
	DIVTR"�ϴ��ļ����ͣ�","�����ϴ��Ŀ������ص����ͣ�ÿ���ַ��á�|������","<input type='text' class='text' name='info34' size=60 style='WIDTH: 99%';' value='"&Info(34)&"' />",55,1
	DIVTR"�ϴ�ͼƬ���ͣ�","�����ϴ��Ŀ�����ʾ��ͼƬ���ͣ�ÿ���ַ��á�|������","<input type='text' class='text' name='info35' size=60 style='WIDTH: 99%';' value='"&Info(35)&"' />",55,2
	DIVTR"�ļ��ϴ�Ŀ¼��","�������,��ͬʱͨ��FTP�½�Ŀ¼���ƶ�ԭ�����ļ�","<input type='text' class='text' name='info36' size='20'  value='"&Info(36)&"' />",55,1	
	DIVTR"ͷ���ϴ�Ŀ¼��","�������,��ͬʱͨ��FTP�½�Ŀ¼���ƶ�ԭ�����ļ�","<input type='text' class='text' name='info37' size='20' value='"&Info(37)&"' />",55,2
	DIVTR"�ϴ���Ϣ�߿�","�Ƿ�����������ʾ�ϴ�����Ϣ�߿�",GetRadio("info39","��",Info(39),0)&GetRadio("info39","��",Info(39),1),55,2

'40-59�û�ѡ��
	.Write"</div><a name='inc4'></a><div class='mian'><div class='top'>�û�ѡ��<a href='#inc'>��</a></div>"
	DIVTR"�û�ע�᣺","�Ƿ������û�ע�᣿",GetRadio("info40","��",Info(40),0)&GetRadio("info40","��",Info(40),1),40,1
	DIVTR"ע����ˣ�","�û�ע����ʺ��Ƿ�Ҫͨ����˲���ʹ��.",GetRadio("info41","��",Info(41),0)&GetRadio("info41","��",Info(41),1),40,2
	DIVTR"ע���������ƣ�","�Ƿ��趨һ������ֻ��ע��һ���ʺ�.",GetRadio("info42","��",Info(42),0)&GetRadio("info42","��",Info(42),1),40,1
	DIVTR"ע�Ỷӭ���ԣ�","�û�ע����,�Ƿ��Զ�����վ�ڻ�ӭ����.",GetRadio("info43","��",Info(43),0)&GetRadio("info43","��",Info(43),1),40,2
	DIVTR"��ӭ�������ݣ�","�û�ע���꣬�Զ����͵����ԡ�","<textarea rows='3' name='info46' cols='70'>"&Info(46)&"</textarea>",55,1
	DIVTR"�ʼƸ���ǩ����","�Ƿ�����������ʾ�û�ǩ��.",GetRadio("info44","��",Info(44),0)&GetRadio("info44","��",Info(44),1),40,1
	DIVTR"��������ʾ��","���û���������ʱ����ʾ��ʽ.",GetRadio("info45","����/ͼ��",Info(45),0)&GetRadio("info45","���ڵ���",Info(45),1),40,2
	DIVTR"�����û���ҳ������","���������û��б��ÿҳ��ʾ����","<input type='text' class='text' name='info47' size='2' value='"&Info(47)&"' /> ",40,2
	DIVTR"�����̳У�","�趨�ϼ��������Թ����¼�����̳.",GetRadio("info48","��",Info(48),0)&GetRadio("info48","��",Info(48),1),40,1
	DIVTR"����ÿ��ÿ�յ�����������","������Ч��������Ȩ����վ���������ƣ�","<input type='text' class='text' name='info49' size='5' value='"&Info(49)&"' /> ��",40,2
'	DIVTR"����������","�����Ƿ���Խ��������ͽ��Ͳ�����",GetRadio("info50","��",Info(50),0)&GetRadio("info50","��",Info(50),1),40,1
	DIVTR"ɾ�����Ӳ���ѡ�","��ɾ����������ʱ���Ƿ���ʾѡ�",GetRadio("info51","��",Info(51),0)&GetRadio("info51","��",Info(51),1),40,2
	DIVTR"��ֹע����û�����","���ڹ����û�����ͷ�γƺ�,�á�|������","<input type='text' class='text' name='info52' style='WIDTH: 99%';'size='60' value='"&Info(52)&"' />",40,1	
	DIVTR"��̳ͷ�������","�趨��̳�Դ�ͷ�����Ŀ","<input type='text' class='text' name='info53' size='5' value='"&Info(53)&"' /> ��",40,2
	DIVTR"ͷ��Ĭ�Ͽ�ȣ�","ͷ���Ĭ�ϴ�߿��","<input type='text' class='text' name='info54' size='5' value='"&Info(54)&"' /> px",40,1
	DIVTR"ͷ��Ĭ�ϸ߶ȣ�","ͷ���Ĭ�ϴ�߸߶�","<input type='text' class='text' name='info55' size='5' value='"&Info(55)&"' /> px",40,2
	DIVTR"ͷ�����ߴ磺","����ͷ�����߶ȺͿ��","<input type='text' class='text' name='info56' size='5' value='"&Info(56)&"' /> px",40,1
	DIVTR"�ⲿͷ��ͼƬ��","�Ƿ����û�ͷ������ⲿ����ͼƬ��",GetRadio("info57","��ֹ",Info(57),0)&GetRadio("info57","����",Info(57),1),40,2
'60-79��������
	.Write"</div><a name='inc5'></a><div class='mian'><div class='top'>��������<a href='#inc'>��</a></div>"
	DIVTR"����ģʽ��","�趨�����༭��",GetRadio("info60","HTML(ȫ����ģʽ)",Info(60),0)&GetRadio("info60","UBB(���ݽ�ʡģʽ)",Info(60),1),40,1
	DIVTR"�����б�������","�����б�(board.asp)ÿҳ����ʾ����","<input type='text' class='text' name='info61' size='5' value='"&Info(61)&"' />",40,2
	DIVTR"���ӻظ�������","������ʾ(Topic.asp)ÿҳ����ʾ����","<input type='text' class='text' name='info80' size='5' value='"&Info(80)&"' />",40,1
	DIVTR"���Ӵ򿪴��ڣ�","�����б�Ĵ򿪷�ʽ",GetRadio("info69","ԭ����",Info(69),0)&GetRadio("info69","�´���",Info(69),1),40,2
	DIVTR"������׼��","��Ϊ��������Ļظ�����","<input type='text' class='text' name='info62' size='5' value='"&Info(62)&"' />",40,2
	DIVTR"ͶƱ������","�û���ͶƱ����������Ŀ","<input type='text' class='text' name='info63' size='5' value='"&Info(63)&"' />",40,1
	DIVTR"�οͲ鿴������ ��","�Ƿ������ο������������",GetRadio("info64","��",Info(64),0)&GetRadio("info64","��",Info(64),1),40,2
	DIVTR"������ͼ��","�Ƿ�������ʶ��UBBͼƬ��ǩ",GetRadio("info65","��",Info(65),0)&GetRadio("info65","��",Info(65),1),40,1
	DIVTR"����ʶ�����ӣ�","�Ƿ����Զ�ʶ�������ϵ���ַ���ӣ�",GetRadio("info82","��",Info(82),0)&GetRadio("info82","��",Info(82),1),40,2
	DIVTR"����Flash��","�Ƿ�������ʶ��UBB������ǩ",GetRadio("info66","��",Info(66),0)&GetRadio("info66","��",Info(66),1),40,1
	DIVTR"������������","�Ƿ�ʶ��UBB������Ƶ������MP/RM",GetRadio("info67","��",Info(67),0)&GetRadio("info67","��",Info(67),1),40,2
	DIVTR"����������룺","�Ƿ���ʶ�����ת����ǩ",GetRadio("info68","��",Info(68),0)&GetRadio("info68","��",Info(68),1),40,1
	DIVTR"������_�ظ��ɼ���","�Ƿ�������ֻ�лظ�����ɼ�����������",GetRadio("info70","��",Info(70),0)&GetRadio("info70","��",Info(70),1),40,1
	DIVTR"������_��Ǯ�ɼ���","�Ƿ�������ﵽָ����Ǯ�����ɼ�����������",GetRadio("info71","��",Info(71),0)&GetRadio("info71","��",Info(71),1),55,2
	DIVTR"������_���ֿɼ���","�Ƿ�������ﵽָ�����ֿɼ�����������",GetRadio("info72","��",Info(72),0)&GetRadio("info72","��",Info(72),1),55,1
	DIVTR"������_���ڿɼ���","�Ƿ���������ָ�����ں�ɼ�����������",GetRadio("info73","��",Info(73),0)&GetRadio("info73","��",Info(73),1),40,2
	DIVTR"������_�Ա�ɼ���","�Ƿ�������ָ���û��Ա�ɼ�����������",GetRadio("info74","��",Info(74),0)&GetRadio("info74","��",Info(74),1),40,1
	DIVTR"������_��½�ɼ���","�Ƿ�������ֻ���ڵ�½��ɼ�����������",GetRadio("info75","��",Info(75),0)&GetRadio("info75","��",Info(75),1),40,2
	DIVTR"������_ָ�����ߣ�","�Ƿ�������ֻ��ָ����Ա�ɼ�����������",GetRadio("info76","��",Info(76),0)&GetRadio("info76","��",Info(76),1),40,1
	DIVTR"������_���ѹۿ���","�Ƿ������������ȡ��Ǯ����������",GetRadio("info77","��",Info(77),0)&GetRadio("info77","��",Info(77),1),40,2
	DIVTR"��������ʾ��","�Ƿ����û���������ʾ��������Ϣ��",GetRadio("info78","��",Info(78),0)&GetRadio("info78","��",Info(78),1),40,1
	DIVTR"���ӹ������֣�","�����ַ����˺��ñ���*������<br>ÿ���ַ����á�|������","<br /><input type='text' class='text' name='info79' style='WIDTH: 99%';' value='"&Info(79)&"' />",55,2	
	DIVTR"���ظ���ʾ��","���û��ظ������ڸ�������ʾ��",GetRadio("info81","����ı���",Info(81),0)&GetRadio("info81","�ظ�������",Info(81),1),40,1
'90-��̳����
	.Write"</div><a name='inc6'></a><div class='mian'><div class='top'>��̳��Դ����<a href='#inc'>��</a></div>"
	DIVTR"��Դ������","��̳��������������ֵ�����Ը�����̳����Ҫ�ĳ��������ơ�<br>���磺���԰ѡ����֡���Ϊ����������","<input type='text' class='text' name='info120' size='20' value='"&Info(120)&"' /> Ĭ�����ƣ���Ǯ<br /><input type='text' class='text' name='info121' size='20' value='"&Info(121)&"' /> Ĭ�����ƣ�����<br /><input type='text' class='text' name='info122' size='20' value='"&Info(122)&"' /> Ĭ�����ƣ���Ϸ��",58,1
	DIVTR"���ö�������","������Ϊ���ö���������ߵĽ���������������Ӧ��Դ",Info(120)&"��<input type='text' class='text' name='info90' size='5' value='"&Info(90)&"' /> "&Info(121)&"��<input type='text' class='text' name='info91' size='5' value='"&Info(91)&"' /> "&Info(122)&"��<input type='text' class='text' name='info92' size='5' value='"&Info(92)&"' />",58,2
	DIVTR"���ö�������","������Ϊ���ö���������ߵĽ���������������Ӧ��Դ",Info(120)&"��<input type='text' class='text' name='info93' size='5' value='"&Info(93)&"' /> "&Info(121)&"��<input type='text' class='text' name='info94' size='5' value='"&Info(94)&"' /> "&Info(122)&"��<input type='text' class='text' name='info95' size='5' value='"&Info(95)&"' />",58,1
	DIVTR"�ö�������","������Ϊ�ö���������ߵĽ���������������Ӧ��Դ",Info(120)&"��<input type='text' class='text' name='info96' size='5' value='"&Info(96)&"' /> "&Info(121)&"��<input type='text' class='text' name='info97' size='5' value='"&Info(97)&"' /> "&Info(122)&"��<input type='text' class='text' name='info98' size='5' value='"&Info(98)&"' />",58,2
	DIVTR"����������","������Ϊ������������ߵĽ���������������Ӧ��Դ",Info(120)&"��<input type='text' class='text' name='info99' size='5' value='"&Info(99)&"' /> "&Info(121)&"��<input type='text' class='text' name='info100' size='5' value='"&Info(100)&"' /> "&Info(122)&"��<input type='text' class='text' name='info101' size='5' value='"&Info(101)&"' />",58,1
	DIVTR"�������⽱����","�û������������Ľ���",Info(120)&"��<input type='text' class='text' name='info102' size='5' value='"&Info(102)&"' /> "&Info(121)&"��<input type='text' class='text' name='info103' size='5' value='"&Info(103)&"' /> "&Info(122)&"��<input type='text' class='text' name='info104' size='5' value='"&Info(104)&"' />",40,2
	DIVTR"����ظ�������","�û�����ظ����Ľ���",Info(120)&"��<input type='text' class='text' name='info105' size='5' value='"&Info(105)&"' /> "&Info(121)&"��<input type='text' class='text' name='info106' size='5' value='"&Info(106)&"' /> "&Info(122)&"��<input type='text' class='text' name='info107' size='5' value='"&Info(107)&"' />",40,1
	DIVTR"ɾ���ͷ���","�����ӱ�ɾ��ʱ�����ߵ�Ĭ�ϳͷ�",Info(120)&"��<input type='text' class='text' name='info108' size='5' value='"&Info(108)&"' /> "&Info(121)&"��<input type='text' class='text' name='info109' size='5' value='"&Info(109)&"' /> "&Info(122)&"��<input type='text' class='text' name='info110' size='5' value='"&Info(110)&"' />",58,2
	DIVTR"�ظ����⽱����","ÿ�λظ�ͬʱ���������ߵĽ���",Info(120)&"��<input type='text' class='text' name='info111' size='5' value='"&Info(111)&"' /> ",58,1
	DIVTR"�����ַ��ٲ�������","�趨С�ڴ��ַ�����������ڽ���","�����ַ�����<input type='text' class='text' name='info112' size='5' value='"&Info(112)&"' /> ",58,2
	DIVTR"�����շѣ�","���û���������ʱ��ȡ����","<input type='text' class='text' name='info123' size='5' value='"&Info(123)&"' /> ��Ǯ",40,1
	DIVTR"��������ޣ�","�趨������ʱ���н��ͷ�����������޶�",Info(120)&"��<input type='text' class='text' name='info113' size='5' value='"&Info(113)&"' /> "&Info(121)&"��<input type='text' class='text' name='info114' size='5' value='"&Info(114)&"' /> "&Info(122)&"��<input type='text' class='text' name='info115' size='5' value='"&Info(115)&"' />",58,1
	DIVTR""&Info(121)&"���ʣ�",""&Info(121)&"�Ļ���","1000��"&Info(120)&" = <input type='text' class='text' name='info116' size='5' value='"&Info(116)&"' /> ��"&Info(121)&"",40,2
	DIVTR""&Info(122)&"���ʣ�",""&Info(122)&"�Ļ���","1000��"&Info(120)&" = <input type='text' class='text' name='info117' size='5' value='"&Info(117)&"' /> ��"&Info(122)&"",40,1
	DIVTR"������н��","����ÿ�µĹ���","<input type='text' class='text' name='info118' size='5' value='"&Info(118)&"' /> "&Info(120)&"",40,2
	DIVTR"�������ʣ�","�û����������е�"&Info(120)&"ÿ������","<input type='text' class='text' name='info119' size='5' value='"&Info(119)&"' /> %",40,1

	.Write"<div class='bottom'><input type='submit' class='button' value='ȷ���޸�'><input type='reset' class='button' value='�� ��'><a href='#inc'>��</a></div></div></form>"
	End with
End Sub


Sub AddMenu
	Dim ParenID,S,Rs
	ParenID=Request("ParenID")
	Response.Write GoForm("SaveMenu")
	Response.Write"<div class='mian'><div class='top'>�����̳�����˵�</div>"
	DIVTR"���ƣ�","","<input type='text' class='text' name='MenuName' size='20' /> *",25,1
	DIVTR"�����ļ���","","<input type='text' class='text' name='MenuUrl' size='35' />(����д���·��,���������ӡ�)",25,2
	DIVTR"�����˵���","",MenuSelect(ParenID),25,1
	DIVTR"��ʾ�ɼ���","","<select name='Show'><option value='0' selected>ȫ���ɼ�</option><option value='1'>ֻ�л�Ա�ɼ�</option><option value='2'>ֻ���οͿɼ�</option><option value='3'>���ɼ�(����)</option></select>",25,2
	DIVTR"�򿪷�ʽ��","","<select name='Target'><option value='0' selected>ԭ����</option><option value='1'>�´���</option></select>",25,1
	Response.Write"<div class='bottom'><input type='submit' value=' �� �� '>&nbsp;&nbsp;<input type='reset' value=' �� �� '></div></div></form>"
End Sub

Sub EditMenu
	Dim ID,Rs,S
	ID=request.querystring("ID")
	Set Rs=BBS.Execute("Select name,Url,Show,Flag,ParenID,Target From [Menu] where ID="&ID&"")
	If Rs.Eof Then Goback"","��¼������"
	Response.Write GoForm("SaveMenu")
	If Rs(3)>0 Then S="ϵͳ<input name='Flag' type='hidden' value='"&Rs(3)&"' />" Else S="��ͨ"
	Response.Write"<div class='mian'><div class='top'>�޸���̳"&S&"�˵�</div>"
	DIVTR"���ƣ�","","<input name='ID' type='hidden' value='"&ID&"' /><input name='MenuName' type='text' class='text' value='"&Rs(0)&"' size='20'> *",25,1
	If Rs(3)>0 Then
		S=Rs(1)
	Else
		S="<input name='MenuUrl' type='text' class='text' value='"&Rs(1)&"' size='38' />(����д���·��,���������ӡ�)"
	End If
	DIVTR"�����ļ���","",S,25,1
	If Rs(3)<>8 Then
	DIVTR"�����˵���","",MenuSelect(Rs(4)),25,1
	DIVTR"�򿪴��ڣ�","",GetRadio("Target","ԭ����",Rs(5),0)&GetRadio("Target","�´���",Rs(5),1),25,1
	End If
	DIVTR"��ʾ�ɼ���","",GetRadio("Show","ȫ���ɼ�",Rs(2),0)&GetRadio("Show","ֻ�л�Ա�ɼ�",Rs(2),1)&GetRadio("Show","ֻ���οͿɼ�",Rs(2),2)&GetRadio("Show","���ɼ�(����)",Rs(2),3),25,1
	Response.Write"<div class='bottom'><input type='submit' value=' �� �� '>&nbsp; <input type='reset' value=' �� �� '></div></div></form>"
	Rs.Close
End Sub

Sub Menu
	Dim Showmood,Sql,Rs1,Subs,I,S
	With Response
	Showmood=Request("Showmood")
	.Write GoForm("MenuOrder")&"<div class='mian'><div class='top'>��̳�˵�</div><div class='divth' style=';padding:5px'><div style='FLOAT: right;'>�鿴��ʽ��<a href='?Action=Menu&Showmood=2'>�οͲ˵�</a> <a href='?Action=Menu&Showmood=1'>��Ա�˵�</a> <a href='?Action=Menu'>��ʾȫ��</a></div><div>��<a href='?Action=AddMenu&ParenID=0'>"&IconA&"��Ӳ˵�</a>�� ��<a href='Admin_Confirm.asp?action=setjsmenu'>���ɲ˵�</a>��</div></div>"
    .Write"<table class='Stable' border='0' cellpadding='3' cellspacing='0'><tr><th width='50px'>����</th><th width='20%'>����</th><th width='30%'>�����ļ�</th><th width='55px'>��ʾ</th><th width='30px'>����</th><th>����</th></tr>"
	Sql="Select ID,Name,Url,show,orders,flag From [Menu] where "
	If Showmood="" Then
		S="ParenID=0 order by orders"
	Else
		S="ParenID=0 and (Show="&Showmood&" or Show=0) order by orders"
	End If
	Set Rs=BBS.Execute(Sql&S)
	Do while Not Rs.eof
	.Write"<tr><td><input name='Orders' type='text' class='text' value='"&Rs(4)&"' size='2'><input name='ID' type='hidden' value='"&Rs(0)&"'></td>"
    .Write"<td>"&Rs(1)&"</td>"
    .Write"<td>"
	If Rs(2)<>"" Then .Write"<a href='"&Rs(2)&"' target='_blank'>"&Rs(2)&"</a>" Else .Write "&nbsp;"
	.Write"</td><td>"&MenuShow(Rs(3))&"</td><td>"
	If Rs(5)=8 Then
		.Write"<font color=red>���</font>"
	ElseIf Rs(5)>0 Then
		.Write"ϵͳ"
	Else
		.Write"��ͨ"
	End If
	.Write"</td><td><a href='?Action=EditMenu&ID="&Rs(0)&"'>"&IconE&"�༭</a> "
	Subs=BBS.Execute("Select Count(*) From [Menu] where parenID="&Rs(0))(0)
	If Rs(5)=0 then
	.Write"<a href=""javascript:"
	If Subs>0 Then
		.Write"alert('�ò˵���������Ŀ������ɾ���������Ƴ����µ������˵���Ŀ��')"">"
	Else
        .Write"checkclick('ɾ���󽫲��ָܻ�����ȷ��Ҫɾ����','Admin_Confirm.asp?Action=DelMenu&ID="&Rs(0)&"')"" >"
	End If 
	.Write IconD&"ɾ��</a> "
	End IF
	If Rs(5)<>8 Then .Write "<a href='?Action=AddMenu&ParenID="&Rs(0)&"'>"&IconA&"���������</a>"
	.Write"</td></tr>"
	'���˵�-ֻ��
	If Rs(5)=8 Then
		Set Rs1=BBS.Execute("Select SkinID,SkinName,IsDefault,Ismode,Pass,remark From [Skins] Order By SkinID Asc")
		Do while not Rs1.eof
			.Write"<tr><td>��</td>"
			.Write"<td>"&Rs1(1)&"</td><td>&nbsp;"
			.Write"</td><td>"&MenuShow(Rs(3))&"</td><td>"
			.Write"���</td><td><a href=""javascript:checkclick('�˲˵�Ϊϵͳ������ڷ������б༭��\n����Ҫת���������б༭��','Admin_Template.asp')"">"&IconE&"�༭</a></td></tr>"
		Rs1.movenext
		Loop
		Rs1.Close
	End If
	'�����˵�
	If Subs>0 Then
		If ShowMood="" Then
			S="parenID="&Rs(0)&" order by orders"
		Else
			S="parenID="&Rs(0)&" and (Show="&showmood&" or Show=0) order by orders"
		End If
		Set Rs1=BBS.Execute(Sql&S)
		Do while Not Rs1.eof
			.Write"<tr><td>��<input name='Orders' type='text' class='text' value='"&Rs1(4)&"' size='2'><input name='ID' type='hidden' value='"&Rs1(0)&"'></td>"
			.Write"<td>"&Rs1(1)&"</td><td>"
			If Rs1(2)<>"" Then .Write"<a href='"&Rs1(2)&"' target=_blank>"&Rs1(2)&"</a>" Else .Write "&nbsp;"
			.Write"</td><td>"&MenuShow(Rs1(3))&"</td><td>"
			If Rs1(5)>0 Then
				.Write"ϵͳ"
			Else
				.Write"��ͨ"
			End If
			.Write"</td><td><a href='?Action=EditMenu&ID="&Rs1(0)&"'>"&IconE&"�༭</a> "
			If Rs1(5)=0 then
			.Write"<a href=""javascript:checkclick('ɾ���󽫲��ָܻ�����ȷ��Ҫɾ����','Admin_Confirm.asp?Action=DelMenu&ID="&Rs1(0)&"')"">"&IconD&"ɾ��</a></td></tr>"
			End If
		Rs1.movenext
		Loop
		Rs1.Close
	End If
	Rs.Movenext
	Loop
	Rs.Close
	Set Rs1=nothing
	.Write"</table><div class='bottom'><input type='submit' class='button' value='��������' /></div></div></form>"
	End With
End Sub

Function MenuShow(Show)
	Select case Show
	case "1"
	MenuShow="ֻ�л�Ա"
	Case "2"
	MenuShow="ֻ���ο�"
	Case "3"
	MenuShow="<font color=blue>����ʾ</font>"
	Case else
	MenuShow="ȫ��ʾ"
	End Select
End Function
Function MenuSelect(parenID)
	Dim mRs,Temp
	Temp="<select name='ParenID'><option value='0'>..��Ϊ�˵�����</option>"
    Set mRs=BBS.Execute("Select ID,Name,Url,show,orders From [Menu] where ParenID=0 And Flag<>8 order by orders")
	Do while Not mRs.eof
        Temp=Temp&"<option value='"&mRs(0)&"'"
		If int(ParenID)=mRs(0) Then Temp=Temp&" selected "
		Temp=Temp&">"&mRs(1)&"</option>"
        mRs.movenext
	Loop
	mRs.Close
	Set mRs=nothing
	Temp=Temp&"</select>"
	MenuSelect=Temp
End Function

Sub ConfigData
	Dim Temp
	With BBS
	If .Cache.valid("Hits") Then Temp=.Cache.Value("Hits")
	Temp=Int(Temp)
	Response.Write GoForm("UpdateConfigData")&"<div class='mian'><div class='top'>��̳ϵͳ��������</div><div class='divth'>˵����������Ϣһ�㲻�����û��޸ģ���*�ŵ���������̳ʱ���ᱻ�Զ�����</div>"
	DIVTR"��̳��Ա���� ��","��̳ע���û�����","<input type='text' name='usernum' size='20' class='text' value='"&.InfoUpdate(5)&"'> *" ,40,1
	DIVTR"��̳�������� ��","��̳������������","<input type='text' name='allessaynum' size='20' class='text' value='"&.InfoUpdate(0)&"'> *" ,40,2
	DIVTR"��̳�������� ��","��̳������������","<input type='text' name='topicnum' size='20' class='text' value='"&.InfoUpdate(1)&"'> *" ,40,1
	DIVTR"��̳����շ�����","��¼��ʷ��ߵ��շ���","<input type='text' name='maxessaynum' size='20' class='text' value='"&.InfoUpdate(4)&"'> " ,40,2
	DIVTR"�������������","��ʷ���ͬʱ���߼�¼����","<input type='text' name='maxonlinenum' size='20' class='text' value='"&.InfoUpdate(7)&"'> " ,40,1
	DIVTR"���������������ʱ�䣺","��ʷ���ͬʱ���߼�¼�������Ǹ�ʱ��","<input type='text' name='maxonlinetime' size='20' class='text' value='"&.InfoUpdate(8)&"'> (��ʽ��YYYY-M-D H:M:S)" ,40,2
	DIVTR"��ִ̳�д�����","ҳ���²��ļ�����","<input type='text' name='hits' size='20' class='text' value='"&.InfoUpdate(9)+Temp&"'>" ,40,1
	Response.Write"<div class='bottom'><input type='submit' value='ȷ���޸�' class='button'><input class='button' type='reset' value='�� ��'></div></div></form>"
	End With
End Sub

Sub A_E_LockIP
	Dim ID,StartIP,EndIp,Readme,Title
	ID=request.querystring("ID")
	StartIP=request.querystring("IP")
	Readme=request.querystring("Readme")
	Title="IP����"
	If ID<>0 Then
		Set Rs=BBS.execute("Select StartIp,EndIp,Readme,ID From[LockIp] where ID="&ID&"")
		IF Rs.eof Then
			GoBack"","��¼������"
			Exit Sub
		Else
			Title="�޸ķ���IP"
			StartIP=BBS.Fun.IpDeCode(Rs(0))
			EndIp=BBS.Fun.IpDeCode(Rs(1))
			Readme=Rs(2)
		End If
	End If
	Response.Write GoForm("LockIp")&"<div class='mian'><div class='top'>"&Title&"</div><input name='ID' type='hidden' value='"&ID&"' />"
	DIVTR"��ʼIP��","���������д","<input name='StartIp' type='text' class='text' value='"&StartIp&"' />",35,1
	DIVTR"����IP��","��������IPʱ������д","<input name='EndIp' type='text' class='text' value='"&EndIp&"' />",35,1
	DIVTR"���˵����","���255���ַ�","<input name='Readme' type='text' class='text' style='width:90%' value='"&Readme&"' />",35,1
	Response.Write"<div class='bottom'><input type='submit' class='button' value='ȷ ��' /><input type='reset' class='button' value='�� ��' /></div></div></form>"
End Sub

Sub LockIp
	Dim S
	A_E_LockIP()
	Response.Write"<div class='mian'><div class='top'>�Ѿ������IP��¼</div>"
	S="<table border='0' class='Stable' cellpadding='3' cellspacing='0'><tr><th width='35%'>����</th><th width='40%'>˵��</th><th>����</th></tr>"
	Set Rs=BBS.Execute("Select StartIp,EndIp,Readme,Lock,ID From[LockIp] where Lock=1")
	If Rs.eof Then
		Response.Write"<div class='divtr1'>û�з�����¼</div>"
	Else
		Response.Write S
		Do while not Rs.eof
			Response.Write"<tr><td>"&BBS.Fun.IpDeCode(Rs("StartIp"))&" �� "&BBS.Fun.IpDeCode(Rs("EndIp"))&"</td><td>&nbsp;"&Rs("Readme")&"</td><td align='center'><a href=?Action=EditIp&Id="&rs("ID")&">"&IconE&"�޸�<a> <a href=Admin_Confirm.asp?Action=IsLockIp&ID="&rs("ID")&"><img src='Images/icon/lock.gif' align='absmiddle' border='0' /> ���</a></td></tr>"
		Rs.MoveNext
		Loop
		Response.Write"</table>"
	End If
	Rs.Close
	Response.Write"</div>"
	Response.Write"<div class='mian'><div class='top'>δ�����IP��¼</div>"
	Set Rs=BBS.Execute("Select StartIp,EndIp,Readme,Lock,ID From[LockIp] where Lock=0")
	If Rs.eof Then
		Response.Write"<div class='divtr1'>û�м�¼</div>"
	Else
	Response.Write S
	Do while not Rs.eof
		Response.Write"<tr><td>"&BBS.Fun.IpDeCode(Rs("StartIp"))&" �� "&BBS.Fun.IpDeCode(Rs("EndIp"))&"</td><td>&nbsp;"&Rs("Readme")&"</td><td align='center'><a href=Admin_Confirm.asp?Action=IsLockIp&ID="&Rs("ID")&"><img src='Images/icon/lock.gif' border=0 align='absmiddle' /> ����</a> <a href=#this onclick=""checkclick('ɾ���󽫲��ָܻ�����ȷ��Ҫɾ����','Admin_Confirm.asp?Action=DelLockIP&ID="&Rs("ID")&"')"">"&IconD&"ɾ��</td></tr>"
	Rs.MoveNext
	Loop
	Response.Write"</table>"
	End If
	Rs.Close	
	Response.Write"</div>"
End Sub

Sub SqlTable
	Dim AllTable,i
	With Response
	.Write"<div class='mian'><div class='top'>���ݱ����</div><div class='divtr2' style='padding:5px'><b>˵����</b><br />Ĭ�����ݱ�Ĭ��ѡ�е�Ϊ��ǰ��̳��ʹ���������������ݵı�<br>ɾ�����ݱ�ɾ����ͬʱȫ��ɾ�������ݱ���������ӣ���ע�⣡����<br>�������ݱ������������ǳ���ʱ������(Access�汾�û�ÿ������5�����ң�SQL�汾�û�ÿ������25������)���һ�����ݱ�<br>�ϲ����ݱ��ϲ��󣬡�ָ�����ݱ��ᱻɾ�������е����ӻ��ƶ�����Ŀ�����ݱ��У�Ĭ�ϱ�����Ϊ��ָ�����ݱ���</div></div>"
	.Write GoForm("AuteSqlTable")&"<div class='mian'><div class='top'>����Ĭ�����ݱ�</div><table class='Stable' border='0' cellpadding='3' cellspacing='0'>"&_
	"<tr><th width='100px'>���ݱ�</th><th width='100px'>����</th><th width='10%'>Ĭ��</th><th>����</th></tr>"
	AllTable=Split(BBS.BBStable(0),",")
	For i=0 To uBound(AllTable)
		.Write"<tr><td>Bbs"&AllTable(i)&"</td><td>"&BBS.execute("Select Count(*) From[bbs"&AllTable(i)&"]")(0)&"</td><td><input name='Aute' type='radio' value='"&AllTable(i)&"'"
		If BBS.BBStable(1)=AllTable(i) Then
			.Write" checked /></td><td><a onclick=alert('�����ݱ�ΪĬ�����ݱ�����ɾ��Ĭ�ϵ����ݱ�') href='#this'>"
		Else
			.Write" /></td><td><a onclick=""checkclick('ע�⣡ɾ�����������ݱ���������ӣ�\n\nɾ���󽫲��ָܻ�����ȷ��Ҫɾ����','Admin_Confirm.asp?Action=DelSqlTable&ID="&AllTable(i)&"')""  href='#'>"
		End If
		.Write IconD&"ɾ��</a></td></tr>"
	Next
	.Write"</table><div class='bottom'><input type='submit' value='�� ��' class='button' /><input type='reset' value='�� ��'  class='button' /></div></div></form>"
	.Write GoForm("AddSqlTable")&"<div class='mian'><div class='top'>�������ݱ�</div><div class='divtr1' style='padding:5px'>�����ݱ����ƣ�bbs<input type='text' name='TableName' class='text' size='2' value='"&uBound(AllTable)+2&"' ONKEYPRESS='event.returnValue=(event.keyCode >= 48) && (event.keyCode <= 57);' /> (ֻ��д���֣����ܺ����е����ݱ���ͬ��)</div>"
	.Write"<div class='bottom'><input type='submit' value='�� ��' class='button' /></div></div></form>"
	.Write GoForm("SqlTableUnite")&"<div class='mian'><div class='top'>�ϲ����ݱ�</div><div class='divtr1' style='padding:5px'>�����ݱ�<select name='SqlTableID1'><option value='0'>ָ�����ݱ�</option>"&GetSqlTableList&"</select> ���е����Ӻϲ��������ݱ� <select name='SqlTableID2'><option value='0'>Ŀ�����ݱ�</option>"&GetSqlTableList&"</select>�У� </div>"
	.Write"<div class='bottom'><input type='button' value='�� ��' class='button' onclick=""if(confirm('ע�⣡�����󽫲��ָܻ�����ȷ��Ҫ�ϲ����ݱ���'))form.submit()"" /></div></div></form>"
	End With
End Sub
Function GetSqlTableList()
	Dim AllTable,I
	AllTable=Split(BBS.BBStable(0),",")
	For i=0 To uBound(AllTable)
	GetSqlTableList=GetSqlTableList&"<option value='"&AllTable(I)&"'>bbs"&AllTable(I)&"</option>"
	Next
End Function


Sub SpaceSize
	dim fso
	On Error Resume Next
	Set fso=server.CreateObject("Scr"&"ipting"&".Fil"&"eSy"&"stemOb"&"ject") 
		If Err Then
			Goback"","�ռ䲻֧��FOS�ļ���д����"
			err.Clear
			Exit Sub
		End If
	Set fso=nothing
	Response.Write"<div class='mian'><div class='top'>ϵͳ�ռ�ռ�����</div>"
	DIVTR"��̳����ռ�ÿռ䣺","","<img src='Images/icon/hr6.gif' style='margin-top:8px' width='"&drawbar("data")&"' height='10' alt='' /> "&GetSpaceinfo("data"),25,1
	DIVTR"��������ռ�ÿռ䣺","","<img src='Images/icon/hr6.gif' style='margin-top:8px' width='"&drawbar("data_backup")&"' height='10' alt='' /> "&GetSpaceinfo("data_backup"),25,2
	DIVTR"�����ļ�ռ�ÿռ䣺","","<img src='Images/icon/hr3.gif' style='margin-top:8px' width='"&drawbar("i@BBS@")&"' height='10' alt='' /> "&GetSpaceinfo("i@BBS@"),25,1
	DIVTR"Inc Ŀ¼ռ�ÿռ䣺","","<img src='Images/icon/hr3.gif' style='margin-top:8px' width='"&drawbar("inc")&"' height='10' alt='' /> "&GetSpaceinfo("inc"),25,2
	DIVTR"ͼƬĿ¼ռ�ÿռ䣺","","<img src='Images/icon/hr5.gif' style='margin-top:8px' width='"&drawbar("pic")&"' height='10' alt='' /> "&GetSpaceinfo("pic"),25,1
	DIVTR"Ƥ��Ŀ¼ռ�ÿռ䣺","","<img src='Images/icon/hr4.gif' style='margin-top:8px' width='"&drawbar("skins")&"' height='10' alt='' /> "&GetSpaceinfo("skins"),25,2
	DIVTR"�ϴ�ͷ��ռ�ÿռ䣺","","<img src='Images/icon/hr2.gif' style='margin-top:8px' width='"&drawbar("UploadFile/Head")&"' height='10' alt='' /> "&GetSpaceinfo("UploadFile/Head"),25,1
	DIVTR"�ϴ��ļ�ռ�ÿռ䣺","","<img src='Images/icon/hr2.gif' style='margin-top:8px' width='"&drawbar("UploadFile/TopicFile")&"' height=10 alt='' /> "&GetSpaceinfo("UploadFile/TopicFile"),25,2
	Response.Write"<div class='bottom' style='padding:2px'>��̳ռ�ÿռ��ܼƣ�<img src='Images/icon/hr1.gif' width='400' height='10' alt='' style='margin-top:8px' /> "&GetSpaceinfo("i@BBS")&"</div></div>"
End Sub
'2005-12-25��д by suibing
Function GetSpaceInfo(drvpath)
	dim fso,d,size,showsize
	Set fso=server.CreateObject("Scr"&"ipting"&".Fil"&"eSy"&"stemOb"&"ject") 
	If Drvpath="i@BBS" Then
		drvpath=server.mappath("Images")
		drvpath=left(drvpath,(instrrev(drvpath,"\")-1))
		set d=fso.getfolder(drvpath) 
		size=d.size
	ElseIf DrvPath="i@BBS@" Then
		dim fc,f1
		drvpath=server.mappath("Images")
		drvpath=left(drvpath,(instrrev(drvpath,"\")-1))
		set d=fso.getfolder(drvpath)
		set fc=d.Files
		for each f1 in fc
			size=size+f1.size
		next
		Set fc=nothing
	Else
		drvpath=server.mappath(drvpath)
		set d=fso.getfolder(drvpath) 		
		size=d.size
	End If
	set d=nothing
	set fso=nothing
	showsize=size & " Byte" 
	if size>1024 then
	   size=(size\1024)
	   showsize=size & " KB"
	end if
	if size>1024 then
	   size=(size/1024)
	   showsize=formatnumber(size,2) & " MB"		
	end if
	if size>1024 then
	   size=(size/1024)
	   showsize=formatnumber(size,2) & " GB"	   
	end if   
	GetSpaceInfo=showsize
End function
Function Drawbar(drvpath)
	dim fso,drvpathroot,d,size,totalsize,barsize
	set fso=server.CreateObject("Scr"&"ipting"&".Fil"&"eSy"&"stemOb"&"ject")
	drvpathroot=server.mappath("Images")
	drvpathroot=left(drvpathroot,(instrrev(drvpathroot,"\")-1))
	set d=fso.getfolder(drvpathroot)
	totalsize=d.size
	If DrvPath="i@BBS@" then
		Dim fc,f1
		set fc=d.files
		for each f1 in fc
			size=size+f1.size
		next
		Set fc=Nothing
	Else
		drvpath=server.mappath(drvpath)
	 On Error Resume Next		
		set d=fso.getfolder(drvpath)
		size=d.size
	End If
	set d=nothing
	set fso=nothing
	barsize=cint((size/totalsize)*300)
	Drawbar=barsize
End Function 
	
Sub UpdateBbs
	Response.Write"<div class='mian'><div class='top'>��̳�����޸�</div>"&_
	"<div class='divth' style='text-align :left;padding:5px'>ע�������̳�����еĸ������ж����ܷǳ����ķ�������Դ��ʱ��Ҳ���ܺܳ��������ĵȺ�<br>��������ѡ����̳�����������ٵ�ʱ��������� ��������������п�������ʱ��<a href=?Action=BbsInfo>�ر���̳</a>��</div>"&_
	"<div class='divtr1' style='padding:5px'><b>��̳ϵͳ����</b><br>���¼��������������������������������û�������ע���û��ȣ�����ÿ��һ��ʱ������һ�Ρ�<br /><input value='��ʼ����' type='button' class='button' onClick=window.location.href='Admin_Confirm.asp?action=UpdateBbsdate' /></div>"&_
	"<div class='divtr2' style='padding:5px'><b>��̳��������</b><br />���¼�����̳��������������������������������������������ظ��ȣ�����ÿ��һ��ʱ������һ�Ρ�����Ĺ������벻Ҫˢ�º͹رգ�<br /><input value='��ʼ����' type='button' class='button' onClick=window.location.href='Admin_Board.asp?Action=BoardUpdate' /></div>"&_
	"<div class='divtr1' style='padding:5px'><b>��̳��������</b><br />������Ч��������Ч���ӡ���Ч���⡢��Ч���ӡ���ЧͶƱ����Ч���ԡ���Ч�û����ȣ�������̿��ܽ����Ĵ�����Դ�������ڱ����Ͻ��У�����Ĺ������벻Ҫˢ�º͹رգ�<br /><input value='��ʼ����' type='button' class='button' onClick=window.location.href='Admin_Confirm.asp?Action=DelWuiong' /></div>"&_
	"<div class='divtr2' style='padding:5px'><b>�޸���������</b><br />�����������ÿ���������Ļظ����������ظ���Ϣ�ȣ������̳���ӷǳ��࣬������̿��ܽ����Ĵ�����Դ��<br /><input value='��ʼ����' type='button' class='button' onClick=window.location.href='Admin_Confirm.asp?Action=UpdateTopic' /></div>"&_
	"<div class='divtr1' style='padding:5px'><b>�޸��û���Ϣ</b><br />�����������ÿ���û��ĵȼ��������������������ȣ����ע���Ա�ǳ��࣬������̿��ܽ����Ĵ�����Դ��<br /><input value='��ʼ����' type='button' class='button' onClick=window.location.href='Admin_Confirm.asp?action=UpdateAllUser' /></div></div>"
End Sub

Sub A_E_Link
	Dim Title,ID,Orders,Ispic,Pic,BbsName,Admin,Url,Readme,Pass
	pass=1
	Ispic=0
	Title="���"
	ID=Request("ID")
	If ID<>"" Then
		Set Rs=BBS.Execute("Select ID,Orders,IsPic,Pic,BbsName,Admin,Url,Readme,pass From [Link] where ID="&ID&"")
		IF Rs.eof Then
			GoBack"","������̳���˲����ڣ�"
			Exit Sub
		Else
			Title="�޸�"
			Orders=Rs(1)
			Ispic=Rs(2)
			Pic=Rs(3)
			BbsName=Rs(4)
			Admin=Rs(5)
			Url=Rs(6)
			Readme=Rs(7)
			Pass=Rs(8)
		End If
		Rs.close
	End If
	Response.Write GoForm("SaveLink")
	Response.Write"<div class='mian'><div class=top>"&Title&"�޸���̳����</div>"
	DIVTR "��̳���ƣ�","","<input name='ID' value='"&ID&"' type='hidden'><input type='text' class='text' name='bbsname' size='15' value='"&BbsName&"' />",25,1
	DIVTR "��̳��ַ��","","<input type='text' name='url' size='28' class='text' value='"&Url&"' />",25,2
	DIVTR "��̳վ����","","<input type='text' name='admin' size='20' class='text' value='"&Admin&"' />(��������)",25,1
	DIVTR "��̳ͼƬ��","","<input type='text' name='pic' size='38' class='text' value='"&Pic&"'>(��ʹ����ͼƬ����-Ҳ������д{�������д})",25,2
	DIVTR "��̳˵����","","<input type='text' name='Readme' size='38' class='text' value='"&Readme&"'>(��������)",25,1
	DIVTR "ͼƬ��ʾ��","",GetRadio("ispic","��",ispic,0)&GetRadio("ispic","��",ispic,1),25,2
	DIVTR "ͨ����ˣ�","",GetRadio("pass","��",pass,0)&GetRadio("pass","<font color=red>�� </font>",pass,1),25,1
	Response.Write"<div class='bottom'><input type='submit' class='button' value='�� ��'><input type='reset' class='button' value='�� ��'></div></div></form>"
End Sub

Sub Grade()
	Dim Arr_Rs,i,S
	Response.Write GoForm("AllUpdateGrade")&"<input name='Grouping' type='hidden' value='0' /><div class='mian'><div class='top'>�û��ȼ�����</div><table border='0' class='Stable' cellpadding='3' cellspacing='0'><tr><th width='15%'>�ȼ�����</th><th width='10%'>��������</th><th width='30%'>�ȼ�ͼƬ</th><th width='19%'>��־ͼƬ</th><th width='20%'>�������</th></tr>"
	Set Rs=BBS.execute("Select Grouping,ID,GradeName,EssayNum,PIC,Spic,Flag FROM [Grade] where Grouping=0 order by EssayNum")
	do while Not Rs.Eof
	Response.Write"<tr><td align='center'><input name='ID' type='hidden' value='"&Rs(1)&"' />"&_
	"<input class='text' name='GradeName' type='text' size='15' value='"&Rs(2)&"' /></td>"&_
	"<td align='center'><input class='text' name='EssayNum' type='text' size='4' value='"&Rs(3)&"' /></td><td><input class='text' name='Pic' type='text' size='15' value='"&Rs(4)&"' /><img src='Pic/Grade/"&Rs(4)&"' /></td>"&_
	"<td><input class='text' name='Spic' type='text' size='15' value='"&Rs(5)&"' />"
	If len(Rs(5))>3 Then Response.Write"<img src='Pic/Grade/"&Rs(5)&"' /></td>"
	Response.Write "<td>"&IconE&"<a href='?action=A_E_Grade&ID="&Rs(1)&"'>�༭Ȩ��</a> <a href=#this onclick=checkclick('"
	If Rs(3)=0 Then Response.Write"ע�⣺�ȼ��������һ���ȼ��ķ�����Ϊ0\nɾ�����ܻᵼ�²���������\n\n" 
	Response.Write"ɾ������ͬʱ�������ڸõȼ�����û���\n\n��ȷ��Ҫɾ����','Admin_Confirm.asp?action=DelGrade&ID="&Rs(1)&"') >"&IconD&"ɾ��"
	Rs.moveNext
	Loop
	Rs.Close
	Response.Write S&"</table><div class='bottom'><input class='button' value='��������' type='button'  onclick=""if(confirm('���������������������������������û��ĵȼ���\n����û��࣬�����Ĵ�����Դ��\n\nȷ��������'))form.submit()"" /><input type='reset' class='button' value='�� ��'>&nbsp;&nbsp; &nbsp;&nbsp;<input class='button' value='��ӵȼ�' type='button' onclick=window.location.href='?Action=A_E_Grade&Grouping=0' /></div></div></form>"
	Response.Write GoForm("AllUpdateGrade")&"<input name='Grouping' type='hidden' value='1' /><div class='mian'><div class='top'>�Զ����ر�ȼ���</div><table border='0' class='Stable' cellpadding='3' cellspacing='0'><tr><th width='15%'>�ȼ�����</th><th width='30%'>�ȼ�ͼƬ</th><th width='19%'>��־ͼƬ</th><th width='20%'>�������</th></tr>"
	Set Rs=BBS.execute("Select Grouping,ID,GradeName,EssayNum,PIC,Spic,Flag FROM [Grade] where Grouping=1 order by ID")
	do while Not Rs.Eof
	Response.Write"<tr><td align='center'><input name='ID' type='hidden' value='"&Rs(1)&"' />"&_
	"<input class='text' name='GradeName' type='text' size='15' value='"&Rs(2)&"' /></td>"&_
	"<td><input class='text' name='Pic' type='text' size='15' value='"&Rs(4)&"' /><img src='Pic/Grade/"&Rs(4)&"' /></td>"&_
	"<td><input class='text' name='Spic' type='text' size='15' value='"&Rs(5)&"' />"
	If len(Rs(5))>3 Then Response.Write"<img src='Pic/Grade/"&Rs(5)&"' /></td>"
	Response.Write "<td>"&IconE&"<a href='?action=A_E_Grade&ID="&Rs(1)&"'>�༭Ȩ��</a> <a href=#this onclick=checkclick('ɾ������ͬʱ�������ڸõȼ�����û���\n\n��ȷ��Ҫɾ����','Admin_Confirm.asp?action=DelGrade&ID="&Rs(1)&"') >"&IconD&"ɾ��"
	Rs.moveNext
	Loop
	Rs.Close
	Response.Write S&"</table><div class='bottom'><input class='button' value='��������' type='submit' /><input type='reset' class='button' value='�� ��'>&nbsp;&nbsp; &nbsp;&nbsp;<input class='button' value='��ӵȼ�' type='button' onclick=window.location.href='?Action=A_E_Grade&Grouping=1' /></div></div></form>"

	Response.Write GoForm("AllUpdateGrade")&"<input name='Grouping' type='hidden' value='2' /><div class='mian'><div class='top'>ϵͳ�̶��ȼ���</div><table border='0' class='Stable' cellpadding='3' cellspacing='0'><tr><th width='15%'>�ȼ�����</th><th width='30%'>�ȼ�ͼƬ</th><th width='19%'>��־ͼƬ</th><th width='8%'>����</th><th width='12%'>�������</th></tr>"
	Set Rs=BBS.execute("Select Grouping,ID,GradeName,EssayNum,PIC,Spic,Flag FROM [Grade] where Grouping=2 order by Flag")
	do while Not Rs.Eof
	Response.Write"<tr><td align='center'><input name='ID' type='hidden' value='"&Rs(1)&"' />"&_
	"<input class='text' name='GradeName' type='text' size='15' value='"&Rs(2)&"' /></td>"&_
	"<td><input class='text' name='Pic' type='text' size='15' value='"&Rs(4)&"' /><img src='Pic/Grade/"&Rs(4)&"' />"&_
	"<td><input class='text' name='Spic' type='text' size='15' value='"&Rs(5)&"' />"
	If len(Rs(5))>3 Then Response.Write"<img src='Pic/Grade/"&Rs(5)&"' />"
	If Rs(6)=9 Then Response.Write"</td><td align='center'>վ��"
	If Rs(6)=8 Then Response.Write"</td><td align='center'>����"
	If Rs(6)=7 Then Response.Write"</td><td align='center'>����"
	If Rs(6)=4 Then Response.Write"</td><td align='center'>VIP"
	Response.Write "</td><td>"&IconE&"<a href='?action=A_E_Grade&ID="&Rs(1)&"'>�༭Ȩ��</a> "
	Rs.moveNext
	Loop		
	Rs.Close
	Response.Write S&"</table><div class='bottom'><input class='button' value='��������' type='submit' /><input type='reset' class='button' value='�� ��'></div></div></form>"
End Sub

Sub A_E_Grade()
	Dim Title,S,Grouping,ID,GradeName,EssayNum,PIC,Spic,Flag,Strings
	ID=request.querystring("ID")
	Grouping=request.querystring("Grouping")
	If ID<>"" Then
		Set Rs=BBS.execute("Select Grouping,ID,GradeName,EssayNum,PIC,Spic,Flag,Strings FROM [Grade] where ID="&ID)
		If Rs.Eof Then
			Goback"","��¼����":Exit Sub
		Else
			Title="�༭�ȼ���"
			Grouping=Rs(0)
			GradeName=Rs(2)
			EssayNum=Rs(3)
			PIC=Rs(4)
			Spic=Rs(5)
			Flag=Rs(6)
			Strings=Split(Rs(7),"|")
		End IF
		Rs.Close
	Else
		PIC="10.Gif"
		EssayNum=0
		Title="��ӵȼ���"
		Strings=Split("#F00|1|0|32100|0|1|0|0|1|1|100|1|50|16000|1|1|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0","|")
	End If
	
	If Grouping=1 Then
		Title=Title&"(�ر���)"
	ElseIf Grouping=2 Then
		Title=Title&"(ϵͳ�̶�)"
	Else
		Title=Title&"(������������)"
	End If

	Response.Write GoForm("SaveGrade")&"<div class='mian'><div class='top'>"&Title&"</div><input name='ID' type='hidden' value='"&ID&"' /><input name='Grouping' type='hidden' value='"&Grouping&"' />"
	DIVTR"�ȼ����ƣ�","","<input name='GradeName' type='text' class='text' size='15' value='"&GradeName&"' />",25,1
	If Grouping=0 Then DIVTR"����ﵽ������","","<input name='EssayNum' type='text' class='text' size='4' value='"&EssayNum&"' />",25,1
	If Pic<>"" Then S="<br /><img src='Pic/Grade/"&Pic&"' />" Else S=""
	DIVTR"�ȼ�ͼƬ��","ͼƬĿ¼\PIC\Grade\","<input name='Pic' type='text' class='text' size='15' value='"&PIC&"' />"&S,42,1
	If sPic<>"" Then S="<img src='Pic/Grade/"&Spic&"' />" Else S=""
	DIVTR"��ݱ�־ͼƬ��","ͼƬĿ¼\PIC\Grade\","<input name='Spic' type='text' class='text' size='15' value='"&Spic&"' />"&S,42,1
	Response.Write"<div class='divth'><li><b>����Ȩ������</b></li></div><div style='clear: both;'></div>"
	DIVTR"������ʾ������ɫ��","","<input name='S0' class='text' type='text' size='15' maxlength='7' value='"&Strings(0)&"' />",25,1
	DIVTR"�Ƿ�����޸��Լ����ϣ�","",GetRadio("S1","��",Strings(1),0)&GetRadio("S1","��",Strings(1),1),25,2
	DIVTR"�Ƿ�����Զ���ͷ�Σ�","",GetRadio("S2","��",Strings(2),0)&GetRadio("S2","��",Strings(2),1),25,1
	DIVTR"���������ַ�����","","<input name='S3' type='text' class='text' size='15' value='"&Strings(3)&"' />���ַ�(����ܳ���65536)",25,2
	DIVTR"�Ƿ���Է�����Ŀ���⣺","",GetRadio("S4","��",Strings(4),0)&GetRadio("S4","��",Strings(4),1),25,1
	DIVTR"�Ƿ���Բμ�ͶƱ���","",GetRadio("S5","��",Strings(5),0)&GetRadio("S5","��",Strings(5),1),25,2
	DIVTR"�Ƿ���Է���ͶƱ���⣺","",GetRadio("S6","��",Strings(6),0)&GetRadio("S6","��",Strings(6),1),25,1
	DIVTR"�Ƿ�����ϴ���","",GetRadio("S8","��",Strings(8),0)&GetRadio("S8","��",Strings(8),1),25,2
	DIVTR"һ����ϴ�������","","<input name='S9' type='text' class='text' size='15' value='"&Strings(9)&"' />��",25,1
	DIVTR"ÿ���ϴ���С��","","<input name='S10' type='text' class='text' size='15' value='"&Strings(10)&"' />KB",25,2
	DIVTR"�Ƿ�����ϴ�ͷ��","",GetRadio("S11","��",Strings(11),0)&GetRadio("S11","��",Strings(11),1),25,1
	DIVTR"��̳�������������","","<input name='S12' type='text' class='text' size='15' value='"&Strings(12)&"' />��",25,2
	DIVTR"����ÿ�췢���ż��Ĵ�����","","<input name='S7' type='text' class='text' size='15' value='"&Strings(7)&"' />",25,1
	DIVTR"����ÿ�����ַ�����","","<input name='S13' type='text' class='text' size='15' value='"&Strings(13)&"' />���ַ�(����ܳ���65536)",25,2
	DIVTR"�Ƿ����������̳��","",GetRadio("S14","��",Strings(14),0)&GetRadio("S14","��",Strings(14),1),25,1
	DIVTR"�Ƿ���Բ鿴������Ϣ��","",GetRadio("S15","��",Strings(15),0)&GetRadio("S15","��",Strings(15),1),25,2
	DIVTR"����ʱ�����Ʊ༭�Լ����ӣ�","",GetRadio("S16","��",Strings(16),0)&GetRadio("S16","��",Strings(16),1),25,1
	DIVTR"��������ɾ���Լ������ӣ�","",GetRadio("S17","��",Strings(17),0)&GetRadio("S17","��",Strings(17),1),25,2
	Response.Write"<div class='divth'><li><b>����Ȩ������</b> ����ֻ�ܹ��������İ��棬�������ޣ���������ѡ�Ҫ��㿪���������������ĵȼ��飩</li></div><div style='clear: both;'></div>"
	DIVTR"���Ա༭���ӣ�","",GetRadio("S18","��",Strings(18),0)&GetRadio("S18","��",Strings(18),1),25,1
	DIVTR"�༭���������뼣��ѡ�","",GetRadio("S19","��",Strings(19),0)&GetRadio("S19","��",Strings(19),1),25,2
	DIVTR"����ɾ�����ӣ�","",GetRadio("S20","��",Strings(20),0)&GetRadio("S20","��",Strings(20),1),25,1
	DIVTR"�����������ӣ�","",GetRadio("S21","��",Strings(21),0)&GetRadio("S21","��",Strings(21),1),25,2
	DIVTR"�����ƶ����ӣ�","",GetRadio("S22","��",Strings(22),0)&GetRadio("S22","��",Strings(22),1),25,1
	DIVTR"�����������⣺","",GetRadio("S23","��",Strings(23),0)&GetRadio("S23","��",Strings(23),1),25,2
	DIVTR"���Գ������⣺","",GetRadio("S24","��",Strings(24),0)&GetRadio("S24","��",Strings(24),1),25,1
	DIVTR"����(��/��)�ö����⣺","",GetRadio("S25","��",Strings(25),0)&GetRadio("S25","��",Strings(25),1),25,2
	DIVTR"����(��/��)���ö����⣺","",GetRadio("S26","��",Strings(26),0)&GetRadio("S26","��",Strings(26),1),25,1
	DIVTR"����(��/��)���ö����⣺","",GetRadio("S27","��",Strings(27),0)&GetRadio("S27","��",Strings(27),1),25,2
	DIVTR"����(��/��)�������⣺","",GetRadio("S28","��",Strings(28),0)&GetRadio("S28","��",Strings(28),1),25,1
	DIVTR"����(��/��)�������⣺","",GetRadio("S29","��",Strings(29),0)&GetRadio("S29","��",Strings(29),1),25,2
	DIVTR"���Խ�����������������","",GetRadio("S30","��",Strings(30),0)&GetRadio("S30","��",Strings(30),1),25,1
	DIVTR"���Բ���ҪͶƱ�ɲ�ͶƱ���飺","",GetRadio("S31","��",Strings(31),0)&GetRadio("S31","��",Strings(31),1),25,2
	DIVTR"���Ա༭ͶƱ��ѡ�","",GetRadio("S32","��",Strings(32),0)&GetRadio("S32","��",Strings(32),1),25,1
	DIVTR"���Բ������������ƣ�","",GetRadio("S33","��",Strings(33),0)&GetRadio("S33","��",Strings(33),1),25,2
	DIVTR"���Է�����̳���棺","",GetRadio("S34","��",Strings(34),0)&GetRadio("S34","��",Strings(34),1),25,1
	DIVTR"����ɾ��������¼��","",GetRadio("S35","��",Strings(35),0)&GetRadio("S35","��",Strings(35),1),25,2
	Response.Write"<div class='divth'><li><b>�߼�����Ȩ������</b> ��������ѡ��ֻ������Ա����</li></div><div style='clear: both;'></div>"
	DIVTR"���Բ鿴�û�IP��","",GetRadio("S36","��",Strings(36),0)&GetRadio("S36","��",Strings(36),1),25,1
	DIVTR"���Բ鿴��̳��־��","",GetRadio("S37","��",Strings(37),0)&GetRadio("S37","��",Strings(37),1),25,2
	'DIVTR"������������������","",GetRadio("S38","��",Strings(38),0)&GetRadio("S38","��",Strings(38),1),25,2
	Response.Write "<div class='bottom'><input type='submit' class='button' value='�� ��' /><input type='reset' class='button' value='�� ��'></div></div></form>"
End Sub

Sub BoardAdmin
	Dim I,po,ii,Name
	Name=Request("Name")
	Response.Write GoForm("BoardAdmin")
	Response.Write"<div class='mian'><div class='top'>��ɾ��̳����</div>"
	Response.Write"<div class='divtr1' style='padding:3px;'><strong>"&BBS.GetGradeName(0,7)&"</strong>��<input name='Name' type='text' class='text' size='12' value='"&Name&"' /> ������<select size='1' name='Flag'><option value='Add'>���</option><option value='Del'>����</option></select> ������̳��<select size='1' name='BoardID'><option value=''>��ѡ�����İ���</option>"&BBS.BoardIDList(0,0)&"</select> <input type='submit' class='button' value='�� ��'></div></form>"
	Response.Write GoForm("AllBoardAdmin")
	Response.Write"<div class='divtr2' style='padding:3px;'><strong>"&BBS.GetGradeName(0,8)&"</strong>��<input name='Name' type='text' class='text' size='12' value='"&Name&"' /> ������<select size='1' name='Flag'><option value='Add'>���</option><option value='Del'>����</option></select> ������̳��ȫ����̳���� <input type='submit' class='button' value='�� ��'></div></form>"
	Response.Write"</div>"
	Response.Write"<div class='mian'><div class='top'>����"&BBS.GetGradeName(0,8)&"</div>"
	Response.Write"<div class='divtr1' style='padding:3px;'>"
	Set Rs=BBS.Execute("Select Name From [Admin] where boardID=-1")
	Do while Not Rs.eof
	Po=Po&"<a href='Admin_user.asp?action=EditUser&Name="&Rs(0)&"'>"&Rs(0)&"</a> &nbsp; &nbsp;"
	Rs.movenext
	loop
	Rs.close
	Response.Write po&"</div></div>"
	Response.Write"<div class='mian'><div class='top'>����"&BBS.GetGradeName(0,7)&"</div>"
	If Not IsArray(BBS.Board_Rs) Then BBS.GetBoardCache()'��ȡ��黺��
	If IsArray(BBS.Board_Rs) Then
		For i=0 To Ubound(BBS.Board_Rs,2)
			po=""
			For II=1 To BBS.Board_Rs(0,i)
				Po=Po&" <font color=red>�O</Font> "
			Next
			If BBS.Board_Rs(0,i)=0 Then
			DIVTR Po&BBS.Board_Rs(3,i),"","",22,2
			Else
			DIVTR Po&BBS.Board_Rs(3,i),"","<div style='padding:3px;'>"&BBS.Board_Rs(6,i)&"</div>",22,1
			End If
		Next
	End If
	Response.Write"</div>"
End Sub

Sub Faction
	Dim UserNum
	Response.Write"<div class='mian'><div class='top'><a style='FLOAT: right;color:#FFF' href='?action=A_E_Faction'>"&IconA&"��Ӱ���&nbsp;</a>��̳���ɹ���</div><table class='Stable' border='0' cellpadding='3' cellspacing='0'>"&_
	"<tr><th>��������</th><th>������</th><th>����</th><th>����ʱ��</th><th>����</th></tr>"
	Set Rs=BBS.Execute("Select ID,Name,User,BuildDate From [Faction] order by ID desc")
	Do while Not Rs.eof
	UserNum=BBS.Execute("select count(ID) from [User] where Faction='"&Rs(1)&"'")(0)
	Response.Write"<tr><td>"&Rs(1)&"</td><td align='center'>"&Rs(2)&"</td><td align='center'>"&UserNum&"</td><td align='center'>"&Rs(3)&"</td>"
	Response.Write"<td align='center'><a href='?Action=A_E_Faction&ID="&Rs(0)&"'>"&IconE&" �༭</a> <a href=#this onClick=""checkclick('ɾ���󽫲��ָܻ�����ȷ��Ҫɾ����','Admin_Confirm.asp?Action=DelFaction&Name="&Rs(1)&"')"">"&IconD&"ɾ��</a></td></tr>"
	Rs.movenext
	Loop
	Rs.Close
	Response.Write"</table></div>"
End Sub

Sub A_E_Faction
	Dim ID,Name,FullName,Note,User,BuildDate,Title
	Id=Request("ID")
	BuildDate=BBS.NowBbsTime
	Title="��Ӱ���"
	If ID<>"" Then
		Set Rs=BBS.Execute("Select Name,FullName,Note,User,BuildDate From [Faction] where ID="&ID)
		IF Rs.eof Then Goback"","��¼������":Exit Sub
		Name=Rs(0)
		FullName=Rs(1)
		Note=Rs(2)
		User=Rs(3)
		BuildDate=Rs(4)
		Title="�༭����"
		Rs.Close
	End If
	Response.Write GoForm("SaveFaction")
	Response.Write"<div class='mian'><div class='top'>"&Title&"</div>"
	DIVTR "�������ƣ�","","<input name='ID' type='hidden' value='"&ID&"' /><input name='Name' class='text' type='text' size='38' value='"&Name&"' />",22,1
	DIVTR "����ȫ�ƣ�","","<input name='FullName' type='text' value='"&FullName&"' maxlength='150' size='50' class='text' />",22,2
	DIVTR "������ּ��","","<input name='Note' type='text' class='text' value='"&Note&"' size='50' maxlength='250' />",22,1
	DIVTR "������","","<input type='text' name='User' size='10' class='text' value='"&User&"' /> �û��������",22,2
	DIVTR "��������","","<input type='text' name='BuildDate' size='20' class='text' value='"&BuildDate&"' />",22,1
	Response.Write"<div class='bottom'><input type='submit' class='button' value='�� ��' /><input class='button' type='reset' value='�� ��' /></div></div></form>"
End Sub

Sub DelEssay
Response.Write GoForm("DelEssay&Go=Date")
Response.Write "<div class='mian'><div class='top'>ɾ��ָ������ǰ������</div>"
DIVTR "ɾ��������ǰ�����ӣ�","","<input name='DateNum' type='text' class='text' value='365' size='5'> ��",22,1
DIVTR "ѡ�����ڵ���̳���棺","","<select name='BoardID'><option value='0'>���е���̳</option>"&BBS.BoardIDList(0,0)&"</select>",22,1
Response.Write "<div class='divtr2' style='padding:3px;'>˵�����˲�����ɾ��ָ������ǰ�������������ͬʱҲ��������Ļظ���(��Ȼ�����������µĻظ���Ҳ�ᱻɾ��)��</div>"
Response.Write"<div class='bottom'><input type='button' class='button' value='�� ��' onclick=""if(confirm('�˲������ɻָ���ȷ��ɾ����'))form.submit()"" /><input class='button' type='reset' value='�� ��' /></div></div></form>"
Response.Write GoForm("DelEssay&Go=DateNoRe")
Response.Write "<div class='mian'><div class='top'>ɾ��ָ������ǰû�лظ�������</div>"
DIVTR "ɾ��������ǰ�����ӣ�","","<input name='DateNum' type='text' class='text' value='100' size='5'> ��",22,1
DIVTR "ѡ�����ڵ���̳���棺","","<select name='BoardID'><option value='0'>���е���̳</option>"&BBS.BoardIDList(0,0)&"</select>",22,1
Response.Write "<div class='divtr2' style='padding:3px;'>˵�����˲�����ɾ��ָ������ǰû���ٻظ���������ͬʱҲ��������Ļظ�����</div>"
Response.Write"<div class='bottom'><input type='button' class='button' value='�� ��' onclick=""if(confirm('�˲������ɻָ���ȷ��ɾ����'))form.submit()"" /><input class='button' type='reset' value='�� ��' /></div></div></form>"
Response.Write GoForm("DelEssay&Go=User")
Response.Write "<div class='mian'><div class='top'>ɾ��ָ���û�����������</div>"
DIVTR "�������û������ƣ�","","<input name='Name' type='text' value='' class='text' size='20'>",22,1
DIVTR "ѡ�����ڵ���̳���棺","","<select name='BoardID'><option value='0'>���е���̳</option>"&BBS.BoardIDList(0,0)&"</select>",22,1
Response.Write "<div class='divtr2' style='padding:3px;'>˵�����˲�����ɾ��ָ���û����������ӡ�</div>"
Response.Write"<div class='bottom'><input type='button' class='button' value='�� ��' onclick=""if(confirm('�˲������ɻָ���ȷ��ɾ����'))form.submit()"" /><input class='button' type='reset' value='�� ��' /></div></div></form>"
End Sub

Sub DelSms
Response.Write GoForm("DelSms&Go=Date")
Response.Write "<div class='mian'><div class='top'>ɾ��ָ������ǰ����������</div>"
DIVTR "ɾ��������ǰ�����ԣ�","","<input name='DateNum' type='text' class='text' value='60' size='5' /> ��",22,1
Response.Write"<div class='bottom'><input type='button' class='button' value='�� ��' onclick=""if(confirm('�˲������ɻָ���ȷ��ɾ����'))form.submit()"" /><input class='button' type='reset' value='�� ��' /></div></div></form>"

Response.Write GoForm("DelSms&Go=Auto")
Response.Write "<div class='mian'><div class='top'>ɾ���Զ����͵��ż�</div>"
DIVTR "ɾ��������ǰ�Զ����͵��ż���","","<input name='DateNum' type='text' class='text' value='60' size='5' /> ��",22,1
Response.Write"<div class='bottom'><input type='button' class='button' value='�� ��' onclick=""if(confirm('�˲������ɻָ���ȷ��ɾ����'))form.submit()"" /><input class='button' type='reset' value='�� ��' /></div></div></form>"

Response.Write GoForm("DelSms&Go=User")
Response.Write "<div class='mian'><div class='top'>ɾ��ָ���û�����������</div>"
DIVTR "������ָ���û����ƣ�","","<input name='Name' type='text' class='text' value='' size='20' />",22,1
Response.Write"<div class='bottom'><input type='button' class='button' value='�� ��' onclick=""if(confirm('�˲������ɻָ���ȷ��ɾ����'))form.submit()"" /><input class='button' type='reset' value='�� ��' /></div></div></form>"
End Sub

Sub MoveEssay
	Response.Write GoForm("MoveEssay&Go=Date")
	Response.Write "<div class='mian'><div class='top'>��ָ�������ƶ�����</div>"
	DIVTR"�ƶ�������ǰ�����ӣ�","","<input name='DateNum' type='text' class='text' value='100' size='5' /> ��",22,2
	DIVTR"����ԭ�����ڵ���̳��","","<select size='1' name='BoardID1'>"&BBS.BoardIDList(0,0)&"</select>",22,1
	DIVTR"����Ҫ�ƶ�������̳��","","<select size='1' name='BoardID2'>"&BBS.BoardIDList(0,0)&"</select>",22,1
	Response.Write"<div class='bottom'><input type='button' class='button' value='�� ��' onclick=""if(confirm('��ȷ��Ҫ�ƶ�������'))form.submit()"" /><input class='button' type='reset' value='�� ��' /></div></div></form>"
	Response.Write GoForm("MoveEssay&Go=User")
	Response.Write "<div class='mian'><div class='top'>�ƶ�ָ���û�������</div>"
	DIVTR"������ָ�����û�����","","<input name='Name' type='text'  size='20' class='text' />",22,2
	DIVTR"����ԭ�����ڵ���̳��","","<select size='1' name='BoardID1'>"&BBS.BoardIDList(0,0)&"</select>",22,1
	DIVTR"����Ҫ�ƶ�������̳��","","<select size='1' name='BoardID2'>"&BBS.BoardIDList(0,0)&"</select>",22,1
	Response.Write"<div class='bottom'><input type='button' class='button' value='�� ��' onclick=""if(confirm('��ȷ��Ҫ�ƶ�����û���������'))form.submit()"" /><input class='button' type='reset' value='�� ��' /></div></div></form>"
End Sub

Sub TopAdmin
	Dim Flag,bgcolor,I,S
	If Instr(AdminString,",22,")=0 Then
	Showtable"��̨Ȩ��","<li>�㱻��ֹ��������Ա��Ȩ�ޣ���</li><li><a href='Admin_user.asp?Action=AdminOK&Name="&BBS.MyName&"'>ֻ���޸��Լ�������</a></li>"
	Footer()
	Response.End
	End If
	Response.Write "<form method=POST  name=form style='margin:0' action='Admin_Confirm.asp?Action=TopAdmin&Flag=1'>"
	Response.Write "<div class='mian'><div class='top'>�����̳����Ա</div>"
	DIVTR"�û����ƣ�","","<input name='Name' type='text' class='text' size='20'>",22,1
	Response.Write"<div class='bottom'><input type='submit' class='button' value='�� ��' /><input class='button' type='reset' value='�� ��' /></div></div></form>"
	Response.Write "<div class='mian'><div class='top'>"&BBS.GetGradeName(0,9)&"�б�</div>"
	i=0
	Set Rs=BBS.execute("Select Name,BoardID From[Admin] where BoardID=0")
	Do while not Rs.eof
	S="<a href=#this onclick=""checkclick('��ȷ��Ҫȡ�������Ա��ְλ��','Admin_Confirm.asp?action=TopAdmin&name="&Rs(0)&"&Flag=0')"">����ְ��</a>"
	IF Rs(0)=BBS.GetMemor("Admin","AdminName") Then S="<a onclick=alert('���ܶ��Լ���ְ��') href='#'>����ְ��</a>"
	S=S&"<a href='Admin_user.asp?action=AdminOK&Name="&Rs(0)&"'>�����ú�̨Ȩ�ޡ�</a>"
	DIVTR "<a href='admin_User.asp?action=EditUser&Name="&Rs(0)&"'>"&Rs(0)&"</a>","","<div style='padding:3px'>"&S&"</div>",22,2
	Rs.movenext
	Loop
	Rs.Close
	Response.Write"</div>"
End Sub

Sub Clean
	Response.Write GoForm("Clean")
	Response.Write "<div class='mian'><div class='top'>���¿ռ仺��</div>"
	Response.Write "<div class='divtr2' style='padding:3px;'>��̳�����˷��������漼����ʹ��̳���ٶȷɿ죬���������̳���ֲ��ȶ���״̬�����������̳�Ļ��档<br />��̳���������ʵʱ������ƣ�һ������²����������̳�Ļ��档<br />������Ա��Ϣ���û����¼�����½�ȫ�������</div>"
	Response.Write "<div class='bottom'><input type='submit' class='button' value='���±�վȫ������'></div></div></form>"
End Sub

Sub Bank
	Response.Write"<form method=POST  style='margin:0' action='Admin_Confirm.asp?Action=Bank' onSubmit=""ok.disabled=true;ok.value='�������ڴ���-���Եȡ�����'"">"
	Response.Write "<div class='mian'><div class='top'>��̨����</div>"
	DIVTR"�û�Ⱥ��","","<select name='user' style='font-size: 9pt'><option value='' selected></option><option value=1>���������û�</option><option value=7>"&BBS.GetGradeName(0,7)&"</option><option value=8>"&BBS.GetGradeName(0,8)&"</option><option value=9>"&BBS.GetGradeName(0,9)&"</option><option value=10>�����Ŷ�(����+����Ա)</option><option value=4>����Vip�û�</option><option value=0>����ע���û�(����)</option></select>",25,1
	DIVTR"������","","<input name='Flag' type='radio' value='1' checked>���� <input name='Flag'  type='radio' value='0'>����",25,1
	DIVTR BBS.Info(120)&"��","","<input name='Coin' type='text' value='1000' >",25,1
	Response.Write "<div class='bottom'><input  type='submit' class='button' value='ȷ��' name='ok'></div></div></form>"
End Sub


Sub GapAd
	Response.Write GoForm("GapAd")
	Response.Write "<div class='mian'><div class='top'>���������</div>"
	Response.Write "<div class='divth' style='padding:5px;text-align: left;'>˵������Щ���ֹ�潫��������������֮�������ʾ��<br />��ʹ�ü򵥵����ֳ�����html���롣<br />�����������ɾ����ع�档</div>"
	Dim S,I,FSO,OpenFile,TmpStr,ad_num,ad_Tmp,BgColor
	Set FSO = server.CreateObject("Scr"&"ipting"&".Fil"&"eSy"&"stemOb"&"ject")
	Set OpenFile=FSO.OpenTextFile(Server.MapPath("inc/ads.js"))
	tmpstr=OpenFile.Readall
	S=split(tmpstr,chr(13)&chr(10))
	ad_num=replace(S(1),";if(a==0){a=1}","")
	ad_num=Int(replace(ad_num,"a=",""))
	i=0
	for i=1 to ad_num
	 ad_Tmp=replace(S(i+8),"b["&i&"].under=","")
	 ad_Tmp=replace(replace(ad_Tmp,"'",""),"<img src=images/icon/ad_icon.gif align=absmiddle> ","")
	 	DIVTR I&"����ʾЧ����","","<div style='line-height:25px'>"&ad_tmp&"</div>",25,2
		DIVTR "&nbsp;&nbsp;&nbsp;��Ӧ���룺","","<textarea  rows='3'  name=ad_v"&i&">"&ad_tmp&"</textarea>",50,1
	Next
	DIVTR"<span style='color:#F00'>���ӹ�棺</span>","","<textarea  rows='3'  name=ad_v"&ad_num+1&"></textarea>",50,2
	Response.Write"<div class='bottom'><input type='submit' class='button' value='ȷ���޸�' /><input class='button' type='reset' value='�� ��' /></div></div></form>"
	OpenFile.close
	Set FSO=Nothing
End Sub
%>
