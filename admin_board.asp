<!--#include file="Admin_Check.asp"-->
<%
Const MaxDepth=5 '�����������������̳�ȼ���ȣ�Ĭ��Ϊ5����̳
Head()
CheckString "11"
Response.Write "<div class='mian'><div class='top'>��̳�������</div><div class='divth'>��<a href='?'>��̳����</a>����<a href='?Action=AddClass'>���ӷ���</a>����<a href='?Action=AddBoard'>������̳</a>����<a href='?Action=ClassOrders'>��������</a>����<a href='?Action=BoardUpdate'>��̳����</a>����<a href='?Action=BoardUnite'>��̳�ϲ�</a>��</div></div>"
Select Case Request("Action")
Case "AddClass"	
CheckString "12"
AddClass
Case "SaveClass"		:SaveClass
Case "EditClass"		:EditClass
Case "SaveEditClass"	:SaveEditClass
Case "DelClass"			:DelClass
Case "AddBoard"	
CheckString "13"
AddBoard
Case "SaveBoard"		:SaveBoard
Case "EditBoard"		:EditBoard
Case "SaveEditBoard"	:SaveEditBoard
Case "DelBoard"			:DelBoard
Case "ClassOrders"		:ClassOrders
Case "SaveClassOrders"	:SaveClassOrders
Case "BoardUnite"		:BoardUnite
Case "SaveBoardUnite"	:SaveBoardUnite
Case "ClearData"		:ClearData
Case "StartClearData"	:StartClearData
Case "PassUser"			:PassUser
Case "SavePassUser"		:SavePassUser
Case "BoardUpdate"		:BoardUpdate
Case "OrdersTopBoard"	:OrdersTopBoard
Case else
	BoardInfo()
end select
Footer()

Sub BoardInfo
	Dim Brs,i,Install,Temp,II,Po,Strings,BoardTypeName
	Response.Write"<div class='mian'><div class='top'>��̳���</div><table class='Stable' border='0' cellpadding='3' cellspacing='0'><tr><th width='180px'>��̳����</th><th width='70px'>���</th><th>��Ӧ����</th></tr>"
	Set Rs=BBS.execute("Select BoardID,BoardName,ParentID,Depth,Child,Strings from [board] order by Rootid,orders")
	If Rs.Bof Then
	    Response.Write "</table></div>"
		GoBack "","��̳û�з��࣡���� <a href='Admin_Board.asp?Action=AddClass'>��ӷ���</a>"
		Exit Sub
	End If
	Brs=Rs.GetRows(-1)
	Rs.close
	For I=0 To Ubound(Brs,2)
		Temp="<tr><td>"
		Install="<a href='?Action=AddBoard&BoardID="&BRs(0,i)&"'>"&IconA&"�����̳</a>"
		If Brs(3,i)=0 Then'����
			BoardTypeName="����"
			Temp="<tr><td>"
			If Brs(4,i)>0 Then'���������̳
				Temp=Temp&Brs(1,i)&" ("&Brs(4,i)&")"
			Else
				Temp=Temp&Brs(1,i)
			End If
			Install=Install &"<a href='?Action=EditClass&BoardID="&Brs(0,i)&"'>"&IconE&"�༭����</a> "
			If Brs(4,i)>0 Then
				Install=Install &"<a href=""javascript:alert('����ɾ�����÷��ຬ����̳!\n\nҪɾ�����࣬�����Ȱ����µ���̳ɾ�������ߡ�')"">"
			Else
				Install=Install &"<a href=""javascript:checkclick('ɾ���󽫲��ָܻ�����ȷ��Ҫɾ����','?Action=DelClass&BoardID="&Brs(0,i)&"')"">"
			End If
			Install=Install&IconD&"ɾ������</a>"
		Else'����
			Strings=Split(Brs(5,i),"|")
		If Strings(7)="1" Then
			BoardTypeName="������̳"
		ElseIf Strings(6)="1" or Strings(5)="1" Then
			BoardtypeName="������̳"
		ElseIf Strings(9)="1" or Strings(3)="1" Then
			BoardTypeName="������̳"
		Else
			BoardtypeName="��ͨ��̳"
		End If
		If Strings(0)="1" Then BoardtypeName=BoardtypeName&"(��)"
			Po=""
			For II=1 To Brs(3,i)
				Po=Po&"<font color=red>|</Font> "
			Next
			If Brs(4,i)>0 Then'���������̳
				Temp=Temp&Po&Brs(1,i)&" ("&Brs(4,i)&")"
			Else
				Temp=Temp&Po&Brs(1,i)
			End If
			Install=Install &"<a href='?Action=EditBoard&BoardID="&Brs(0,i)&"'>"&IconE&"��������</a>"
			If Brs(4,i)>0 Then
				Install=Install &" <a href=""javascript:alert('����ɾ�����ð��溬������̳!\n\nҪɾ�����棬�����Ȱ����µ�����̳ɾ�������ߡ�')"">"&IconD&"ɾ������</a>"
			Else
				Install=Install &" <a href=""javascript:checkclick('ɾ���󽫲��ָܻ�����ȷ��Ҫɾ����','?Action=DelBoard&BoardID="&Brs(0,i)&"')"">"&IconD&"ɾ����̳</a>"
			End If
				Install=Install &" <a href='?Action=ClearData&BoardID="&BRs(0,i)&"'><img src='images/Icon/recycle.gif' border='0' align='absmiddle' /> ��������</a>"
				Install=Install & " <a href='?Action=OrdersTopBoard&BoardID="&BRs(0,i)&"'><img src='Images/icon/Top.gif' border='0' align='absmiddle' /> ��������</a>"
			If Strings(6)="1" Then
				Install=Install & " <a href='?Action=PassUser&BoardID="&BRs(0,i)&"'><img src='Images/icon/user.gif' border='0' align='absmiddle' /> ��֤�û�</a>"
			End If
		End If
		Response.Write Temp&"</td><td align='center'><span style='color:#888'>"&BoardTypeName&"</span></td><td>"&Install&"</td></tr>"
	Next
	Response.Write"</table></div>"
End Sub

Sub AddClass
	Dim NewBoardID
	Set Rs=BBS.Execute("select Max(BoardID) from [Board]")
	IF Rs.Eof or Rs.Bof Then
		NewBoardID=1
	Else
		NewBoardID=Rs(0)+1
	End If
	If Not isnumeric(NewBoardID) Then NewBoardID=1
	Rs.Close
	Response.Write"<form method=POST style=""margin:0"" action=""?Action=SaveClass""><div class='mian'><div class='top'>��ӷ���</div>"
	DIVTR"�������ƣ�","��̳�ķ�������","<input name=""NewBoardID"" type=""hidden"" value='"&NewBoardID&"' /><input type=""text"" class='text' class='text' name='BoardName' size='30'>",40,1
	DIVTR"��̳�����ʾ��","���ð����������̳�Ƿ��Լ�෽ʽ��ʾ","<input name='s2' type='radio' value='0' checked>��<input name='s2' type='radio' value='1'>��",40,1
	DIVTR"�����ʾ������","������Ϊ�����ʾ��ÿһ����ʾ�ĸ���[һ����ʾ4���Ƚ�����]","<input name='s3' type='text' class='text' size='3' maxlength='2' value='4'>��",40,1
	Response.Write"<div class='bottom'><input type=""submit"" class='button' value=""�� ��""><input type=""reset"" value=""�� ��"" class='button'></div></div></form>"
End Sub

Sub SaveClass
	Dim BoardName,NewBoardID,MaxRootID,Rs,S1,S2,Temp
	BoardName=Replace(BBS.Fun.GetStr("BoardName"),",","&#44")
	NewBoardID=BBS.Fun.GetStr("NewBoardID")
	S1=BBS.Fun.GetStr("s1")
	S2=BBS.Fun.GetStr("s2")
	IF BoardName="" Or Not isnumeric(NewBoardID) Then
		GoBack"","":Exit Sub
	Else
		Set Rs=BBS.Execute("select BoardID from [Board] where BoardID="&NewBoardID)
		If Not (rs.eof and rs.bof) then
			GoBack"�ڲ�ϵͳ����","����ָ���ͱ����̳һ������ţ�������ܽ�������⣬�뵽BBS�ٷ���̳Ѱ�������"
			Exit Sub
		End if
		Rs.Close
		Set Rs=BBS.Execute("Select Max(RootID) From [Board]")
		MaxRootID=Rs(0)+1
		If isnull(MaxRootID) then MaxRootID=1
		Rs.Close
		BBS.execute("Insert into [Board](BoardName,BoardID,RootID,Depth,ParentID,Orders,Child,ParentStr,Strings)Values('"&BoardName&"',"&NewBoardID&","&MaxRootID&",0,0,0,0,'0','0|"&s1&"|"&s2&"|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0')")
		BBS.Cache.clean("BoardInfo")
		Temp="�������̳���� <b>"&BoardName&" </b> �ɹ�!"
		BBS.NetLog "������̨_"&Temp
		Suc"",Temp,"?"
	End If
End Sub

Sub EditClass
	Dim BoardID,Rs,Strings
	Set Rs=BBS.Execute("Select BoardName,Strings from[Board] where BoardID="&BBS.BoardID&"")
	If Rs.Eof Then
		GoBack "ϵͳ����","��̳�Ҳ���������࣬�����Ѿ�ɾ���ˡ�":Exit Sub
	End If
	Strings=Split(Rs(1),"|")
	Response.Write"<form method=POST style=""margin:0"" action=""?Action=SaveEditClass""><div class='mian'><div class='top'>�༭����</div>"
	DIVTR"�������ƣ�","�޸���̳���������","<input name=""BoardID"" type=""hidden"" value='"&BBS.BoardID&"' /><input type=""text"" class='text' class='text' name='BoardName' size='30' value='"&Rs(0)&"' />",40,1
	DIVTR"��̳�����ʾ��","���ð����������̳�Ƿ��Լ�෽ʽ��ʾ",GetRadio("s1","��",Strings(1),0)&GetRadio("s1","��",Strings(1),1),40,1
	DIVTR"�����ʾ������","������Ϊ�����ʾ��ÿһ����ʾ�ĸ���[һ����ʾ4���Ƚ�����]","<input name='s2' type='text' class='text' size='3' maxlength='2' value='"&Strings(2)&"'>��",40,1
	Response.Write"<div class='bottom'><input type=""submit"" value=""�� ��"" class='button'><input type=""reset"" value=""�� ��"" class='button'></div></div></form>"
	Rs.Close
End Sub

Sub SaveEditClass
	Dim BoardName,BoardID,S1,S2,Temp
	BoardName=Replace(BBS.Fun.GetStr("BoardName"),",","&#44")
	BoardID=Request.Form("BoardID")
	S1=BBS.Fun.GetStr("s1")
	S2=BBS.Fun.GetStr("s2")
	IF BoardName="" Or Not isnumeric(BoardID) Then
		GoBack"","":Exit Sub
	Else
		If BBS.Execute("select BoardID from [Board] where BoardID="&BoardID).eof then
			GoBack"ϵͳ����","��̳�Ҳ���������࣬�����Ѿ�ɾ���ˡ�":Exit Sub
		End if
		BBS.execute("Update [Board] Set BoardName='"&BoardName&"',Strings='0|"&s1&"|"&s2&"|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0' where BoardID="&BoardID)
		Temp="��̳�������Ƹ�Ϊ <b>"&BoardName&"</b> �ɹ�!"
		BBS.NetLog "������̨_"&Temp
		BBS.Cache.clean("BoardInfo")
		Suc "",Temp,"?"
	End If
End Sub

Sub DelClass
	Dim Temp
	IF BBS.Execute("Select Count(BoardID) From[Board] where ParentID="&BBS.BoardID)(0)=0 Then
		BBS.Execute("Delete From[Board] where ParentID=0 And BoardID="&BBS.BoardID)
		BBS.Cache.clean("BoardInfo")
		Temp="ɾ����̳����ɹ�!"
		BBS.NetLog "������̨_"&Temp
		Suc"",Temp,"?"
	End If
End Sub

Sub DelBoard
	Dim AllTable,I,II,Depth,ParentID,RootID,Orders,Temp
	Set Rs=BBS.Execute("Select Depth,ParentID,RootID,Orders,Child From[Board] where BoardID="&BBS.BoardID)
	If Rs.Eof Then 
		Goback"","�����ڣ���̳�����Ѿ�ɾ���� !"
		Exit Sub
	ElseIf Rs(4)>0 Then
		Goback"","����̳����������̳������ɾ�� !"
		Exit sub
	Else
		Depth=Rs(0)
		ParentID=Rs(1)
		RootID=Rs(2)
		Orders=Rs(3)
	End If
	Rs.Close
	AllTable=Split(BBS.BBStable(0),",")
	For i=0 To uBound(AllTable)
		BBS.Execute("Delete from [Bbs"&AllTable(i)&"] where BoardID="&BBS.BoardID&"")
	Next

	Set Rs=BBS.Execute("Select TopicID From[Topic] where BoardID="&BBS.BoardID&" And IsVote=1 ")
	Do While Not Rs.Eof
		BBS.Execute("Delete from [TopicVote] where TopicID="&Rs(0))
		BBS.Execute("Delete from [TopicVoteUser] where TopicID="&Rs(0))
	Rs.MoveNext
	Loop
	Rs.Close
	'ɾ�������¼
	BBS.Execute("Delete From[Topic] where BoardID="&BBS.BoardID)
	BBS.Execute("Delete From[Board] where BoardID="&BBS.BoardID)
	'�����丸��İ�����
	BBS.Execute("update [Board] set Child=child-1 where BoardID="&ParentID)
	'�����丸���������������
	BBS.Execute("Update [Board] Set Orders=Orders-1 where RootID="&RootID&" and Orders>"&Orders)
	'������������̳����
		For II=1 to Depth
			'�õ��丸��ĸ���İ���ID
			Set rs=BBS.Execute("select ParentID from [Board] where BoardID="&ParentID)
			if not (rs.eof and rs.bof) then
				ParentID=rs(0)
				'�����丸��ĸ��������
				BBS.Execute("update [Board] set child=child-1 where boardid="&ParentID)
			End IF
			Rs.Close
		Next
	BBS.Cache.clean("BoardInfo")
	BBS.Cache.clean("Board"&BBS.BoardID)
	Temp="�ɹ���ɾ����̳���� (��������̳����������)!"
	BBS.NetLog "������̨_"&Temp
	Suc"",Temp,"Admin_Board.asp"
End Sub


Sub AddBoard
	If BBS.execute("Select BoardID from [Board] where Depth=0").Eof Then
		GoBack"","û�з��಻�������̳������ <a href=Admin_Board.asp?Action=AddClass>��ӷ���</a>"
		Exit Sub
	End if
	Response.Write"<form method=POST style='margin:0' action='?Action=SaveBoard'>"
	Response.Write"<div class='mian'><div class='top'>��̳��� </div>"
	DIVTR"���ڷ������̳��","ѡ��Ҫ�����Ǹ�������Ǹ���̳","<select name='ParentID'>"&BBS.BoardIDList(BBS.BoardID,20)&"</select>*",40,1
	DIVTR"��̳���ƣ�","��̳���������","<input type='text' class='text' name='BoardName' size='30' />*",40,2
	DIVTR"��־ͼƬ��","��̳����Logo��ַ��Ϊ����ҳ����һ��Ҫ��д","<input type='text' class='text' name='BoardImg' size='30' />*",40,1
	DIVTR"��̳���ܣ�","��̳��������","<textarea rows='3' name='Introduce'  cols='60'></textarea>*",58,2
	DIVTR"��ΪС�ࣺ","�Ƿ�����Ϊ�࣬���ú�ð��治�ܷ���","<input name='s0' type='radio' value='0' checked>��<input name='s0' type='radio' value='1'>��",40,1
	DIVTR"��̳�����ʾ��","���ð����������̳�Ƿ��Լ�෽ʽ��ʾ","<input name='s1' type='radio' value='0' checked>��<input name='s1' type='radio' value='1'>��",40,2
	DIVTR"�����ʾ������","������Ϊ�����ʾ��ÿһ����ʾ�ĸ���[һ����ʾ4���Ƚ�����]","<input name='s2' type='text' class='text' size='3' value='4' maxlength='2'>��",40,1
	DIVTR"��̳���ͣ�","������̳�����ͣ����Զ�ѡ","<input type='checkbox' name='s3' value='1' />��Ա��ֻ�л�Ա����������ӣ� <br /><input type='checkbox' name='s4' value='1' />ֻ��������������ӣ���ֻ��վ�������桢�����ܷ�����<br /><input type='checkbox' name='s5' value='1' />VIP��ֻ��vip�û����ܽ��룩<br /><input type='checkbox' name='s6' value='1' />��֤��ֻ��ͨ����֤���û����ܽ��룩",90,2
	DIVTR"������̳��","��̳����վ����һ�ɲ��ý���","<input name='s7' type='radio' value='0' checked>����<input name='s7' type='radio' value='1'>����",40,1
	DIVTR"�������ƣ�","�û��ﵽ��Щ��Դ����Խ���","������<input name='s10' type='text' class='text' value='0' size='6' /><br>"&BBS.Info(121)&"��<input name='s11' type='text' class='text' value='0' size='6' /><br>"&BBS.Info(120)&"��<input name='s12' type='text' class='text' value='0' size='6' /><br>"&BBS.Info(122)&"��<input name='s13' type='text' class='text' value='0' size='6' />",120,2
	DIVTR"�ϴ����ã�","����̳ϵͳĬ�Ͻ�ֹ�ϴ��󣬽��������á�","<input type='radio' name='s14' value='0' />��ֹ <input type='radio' name='s14' value='1' checked />ȫ����Ա <input type='radio' name='s14' value='2' />ֻ��VIP��Ա",40,1
	Response.Write"<div class='bottom'><input type='submit' value=' �� �� ' class='button' /><input type='reset' value=' �� �� ' class='button' /></div></div></form>"
End Sub

Sub SaveBoard
	Dim Strings(14),BoardName,Introduce,BoardImg,ParentID,NewBoardID,RootID,Depth,Child,Orders,ParentStr,I
	BoardName=Replace(BBS.Fun.GetStr("BoardName"),",","&#44")
	Introduce=BBS.Fun.GetStr("Introduce")
	BoardImg=BBS.Fun.GetStr("BoardImg")
	ParentID=BBS.Fun.GetStr("ParentID")
	For i =0 to 14
		Strings(i)=Request.Form("s"&i)
		If Strings(i)="" Then Strings(i)=0
		If Not Isnumeric(strings(i)) Then GoBack"","һЩ��Ŀ������������д��":Exit Sub
	Next
	'�������������Դ
	For I=10 To 13
		If Int(Strings(I))>0 Then
			Strings(9)=1
			Exit For
		End If
	Next
	If Not isnumeric(ParentID) or BoardName="" Or Introduce="" Then
		GoBack"","":Exit Sub
	End If
	
	Set Rs=BBS.Execute("select Max(BoardID) from [Board]")
	IF Rs.Eof or Rs.Bof Then
		GoBack"","û�з��಻�������̳������<a href='Admin_Board.asp?Action=AddClass'>��ӷ���</a>"
		Exit Sub
	Else
		NewBoardID=Rs(0)+1
	End If
	Rs.Close
	Set Rs=BBS.execute("Select RootID,Depth,Child,Orders,ParentStr,ParentID From[Board] where BoardID="&ParentID&"")
	IF Rs.Eof or Rs.Bof Then
		GoBack"ϵͳ��ʽ����","û��ָ���������̳��"
		Exit Sub
	End If
	RootID=Rs(0)
	Depth=Rs(1)
	Child=Rs(2)
	Orders=Rs(3)
	ParentStr=Rs(4)
	Rs.Close
	If Depth+1>MaxDepth Then
		GoBack "","���˵���̳��ʵ�����ã�����̳���������ֻ����" & MaxDepth & "����̳��^_^"
		Exit Sub
	End If
	If ParentStr=0 Then
		ParentStr=ParentID
	Else
		ParentStr=ParentStr & "," & ParentID
	End If
	BBS.execute("Insert into [Board](BoardID,BoardName,Introduce,BoardImg,RootID,Depth,ParentID,ParentStr,Orders,Child,Strings,LastReply)Values("&NewBoardID&",'"&BoardName&"','"&Introduce&"','"&BoardImg&"',"&RootID&","&Depth+1&","&ParentID&",'"&ParentStr&"',"&Orders+1&",0,'"&join(Strings,"|")&"','|||||||||')")
	If ParentID<>0 then
	If Depth>0 then
		'���ϼ�������ȴ���0��ʱ��Ҫ�����丸�ࣨ����ĸ��ࣩ�İ��������������
		For i=1 to Depth
			'�����丸�������
			BBS.Execute("update [Board] set Child=Child+1 where BoardID="&parentID)
			'�õ��丸��ĸ���İ���ID
			Set rs=BBS.Execute("select ParentID from [Board] where BoardID="&parentID)
			If not (rs.eof and rs.bof) then
				ParentID=rs(0)
			End if
			Rs.Close
			'��ѭ����������1�������е����һ��ѭ����ʱ��ֱ�ӽ��и���
			If i=depth then
			BBS.Execute("update [Board] set Child=Child+1 where BoardID="&parentID)
			End if
		next
		'���¸ð��������Լ����ڱ���Ҫ��ͬ�ڱ������µİ����������
		BBS.Execute("update [Board] set Orders=orders+1 where RootID="&RootID&" And orders>"&orders)
		BBS.Execute("update [Board] set Orders="&Orders&"+1 where BoardID="&NewBoardID&"")
	Else
		'���ϼ��������Ϊ0��ʱ��ֻҪ�����ϼ����������
		BBS.Execute("update [Board] set child=child+1 where Boardid="&ParentID)
		Set rs=BBS.Execute("select max(Orders) from [Board]")
		BBS.Execute("update [Board] set Orders="&rs(0)&"+1 where BoardID="&NewBoardID )
		Rs.Close
	End if
	End if
	If Strings(6)="1" Then
		Suc"","�ɹ����������̳ <b>"&BoardName&"</b> !<li>����̳Ϊ��֤��̳����ʱֻ����߹���Ա�ܹ����롣<li>�����ͨ�� <a href=?Action=PassUser>����</a> ��Ŀ����ӿ��Խ������̳���û�","Admin_Board.asp"	
	Else
		Suc"","�ɹ����������̳ <b>"&BoardName&"</b> !","Admin_Board.asp"	
	End IF
	BBS.NetLog"������̨_�����̳<b>"&BoardName&"</b>�ɹ�!"
	BBS.Cache.clean("BoardInfo")
	BBS.Cache.clean("Board"&NewBoardID)
End Sub

Sub EditBoard
	Dim BoardName,Strings,Introduce,BoardImg,ParentID
	Dim Temp,Chk
	Set Rs=BBS.execute("Select ParentID,BoardName,Strings,Introduce,BoardImg From[Board] Where BoardID="&BBS.BoardID&"")
	If Rs.eof Then
		GoBack"","�ð��治���ڣ������Ѿ�ɾ����"
		Exit Sub
	Else
		ParentID=Rs(0)
		BoardName=Rs(1)
		Strings=split(Rs(2),"|")
		Introduce=Rs(3)
		BoardImg=Rs(4)
	End IF
	Rs.Close
	Response.Write"<form method=POST style='margin:0' action='?Action=SaveEditBoard'>"
	Response.Write"<div class='mian'><div class='top'>�༭��̳ </div><input name='BoardID' type='hidden' value='"&BBS.BoardID&"'>"
	
	DIVTR"��̳���ƣ�","��̳���������","<input type='text' class='text' name='BoardName' size='30' value='"&BoardName&"' />*",40,1
	DIVTR"���ڷ������̳��","ѡ��Ҫ�����Ǹ�������Ǹ���̳","<select name='ParentID'>"&BBS.BoardIDList(ParentID,20)&"</select>*",40,2
	DIVTR"��־ͼƬ��","��̳����Logo��ַ��Ϊ����ҳ����һ��Ҫ��д","<input type='text' class='text' name='BoardImg' size='30' value='"&BoardImg&"' />*",40,1
	DIVTR"��̳���ܣ�","��̳��������","<textarea rows='3' name='Introduce'  cols='60'>"&Introduce&"</textarea>*",58,2
	DIVTR"��ΪС�ࣺ","�Ƿ�����Ϊ�࣬���ú�ð��治�ܷ���",GetRadio("s0","��",Strings(0),0)&GetRadio("s0","��",Strings(0),1),40,1
	DIVTR"��̳�����ʾ��","���ð����������̳�Ƿ��Լ�෽ʽ��ʾ",GetRadio("s1","��",Strings(1),0)&GetRadio("s1","��",Strings(1),1),40,2
	DIVTR"�����ʾ������","������Ϊ�����ʾ��ÿһ����ʾ�ĸ���[һ����ʾ4���Ƚ�����]","<input name='s2' type='text' class='text' size='3' value='"&Strings(2)&"' maxlength='2'>��",40,1
	If Strings(3)="1" Then Chk="checked" Else Chk=""
	Temp="<input type='checkbox' name='s3' value='1' "&Chk&" />��Ա��ֻ�л�Ա����������ӣ� <br />"
	If Strings(4)="1" Then Chk="checked" Else Chk=""
	Temp=Temp&"<input type='checkbox' name='s4' value='1' "&Chk&" />ֻ��������������ӣ���ֻ��վ�������桢�����ܷ�����<br />"
	If Strings(5)="1" Then Chk="checked" Else Chk=""
	Temp=Temp&"<input type='checkbox' name='s5' value='1' "&Chk&" />VIP��ֻ��vip�û����ܽ��룩<br />"
	If Strings(6)="1" Then Chk="checked" Else Chk=""
	Temp=Temp&"<input type='checkbox' name='s6' value='1' "&Chk&" />��֤��ֻ��ͨ����֤���û����ܽ��룩"
	DIVTR"��̳���ͣ�","������̳�����ͣ����Զ�ѡ",Temp,90,2
	DIVTR"������̳��","��̳����վ����һ�ɲ��ý���",GetRadio("s7","����",Strings(7),0)&GetRadio("s7","����",Strings(7),1),40,1
	DIVTR"�������ƣ�","�û��ﵽ��Щ��Դ����Խ���","������<input name='s10' type='text' class='text' value='"&Strings(10)&"' size='6' /><br>"&BBS.Info(121)&"��<input name='s11' type='text' class='text' value='"&Strings(11)&"' size='6' /><br>"&BBS.Info(120)&"��<input name='s12' type='text' class='text' value='"&Strings(12)&"' size='6' /><br>"&BBS.Info(122)&"��<input name='s13' type='text' class='text' value='"&Strings(13)&"' size='6' />",90,2
	DIVTR"�ϴ����ã�","����̳ϵͳĬ�Ͻ�ֹ�ϴ��󣬽��������á�",GetRadio("s14","��ֹ",Strings(14),0)&GetRadio("s14","ȫ����Ա",Strings(14),1)&GetRadio("s14","ֻ��VIP��Ա",Strings(14),2),40,1
	Response.Write"<div class='bottom'><input type='submit' value='�� ��' class='button' /><input type='reset' value=' �� �� ' class='button' /></div></div></form>"
End Sub

Sub SaveEditBoard
	Dim Strings(14),BoardID,BoardName,Introduce,BoardImg,ParentID,RootID,Depth,Child,Orders,ParentStr,I
	Dim NewParentID,BoardNum,P_Rs
	BoardID=BBS.Fun.GetStr("BoardID")
	BoardName=Replace(BBS.Fun.GetStr("BoardName"),",","&#44")
	Introduce=BBS.Fun.GetStr("Introduce")
	BoardImg=BBS.Fun.GetStr("BoardImg")
	NewParentID=BBS.Fun.GetStr("ParentID")
	For i =0 to 14
		Strings(i)=Request.Form("s"&i)
		If Strings(i)="" Then Strings(i)=0
		If Not Isnumeric(strings(i)) Then GoBack"","һЩ��Ŀ������������д��":Exit Sub
	Next
	'�������������Դ
	For I=10 To 13
		If Int(Strings(I))>0 Then
			Strings(9)=1
			Exit For
		End If
	Next
	
	If Not isnumeric(NewParentID) or BoardName="" Or Introduce="" Then
		GoBack"","":Exit Sub
	ElseIF BoardID=NewParentID Then
		GoBack"","������̳����ָ���Լ���":Exit Sub
	End If
	Set Rs=BBS.execute("Select RootID,Depth,Child,Orders,ParentID,ParentStr From[Board] where BoardID="&BoardID)
	IF Rs.Eof or Rs.Bof Then
		GoBack"ϵͳ����","�ð��治���ڣ������Ѿ�ɾ���ˣ�"
		Exit Sub
	End If
	RootID=Rs(0)
	Depth=Rs(1)
	Child=Rs(2)
	Orders=Rs(3)
	ParentID=Rs(4)
	ParentStr=Rs(5)
	Rs.Close
	If ParentID=0 then
		GoBack"ϵͳ����","���಻������":Exit Sub
	ElseIf Int(NewParentID)<>Int(ParentID) Then
		'�ж���ָ������̳�Ƿ���������̳
		Set Rs=BBS.Execute("select BoardID from [board] where ParentStr like '%"&ParentStr&","&BoardID&"%' and BoardID="&NewParentID)
		if not (Rs.eof and Rs.bof) then
			GoBack"","������ָ���ð������������̳��Ϊ������̳"
			Exit sub
		End if
		Rs.Close
		'�����ѡ�ĸ���
		Set P_rs=BBS.Execute("select * from [board] where Boardid="&NewParentID)
			If P_rs("Depth")+1> MaxDepth Or (Child>0 And P_Rs("Depth")+2>MaxDepth) Then
			GoBack "","����̳���������ֻ����" & MaxDepth & "����̳�������ʹ�ø��༶��̳���뵽BBS�ٷ���̳Ѱ�������"
			P_rs.Close
			Set P_rs=Nothing
			Exit Sub
		End If		
	End if
	BBS.Execute("Update [Board] Set BoardName='"&BoardName&"',Strings='"&Join(Strings,"|")&"',Introduce='"&Introduce&"',BoardImg='"&BoardImg&"' where BoardID="&BoardID&"")
  If Int(NewParentID)<>Int(ParentID) Then
	'��һ������̳�ƶ�����������̳��
	'�����ָ������̳�������Ϣ
	'�õ�������������
	ParentStr=ParentStr & ","
	BoardNum=BBS.Execute("select count(*) from [Board] where ParentStr like '%"&ParentStr & BoardID&"%'")(0)
	If Isnull(BoardNum) Then BoardNum=1
	'�ڻ���ƶ������İ����������������ָ����̳֮�����̳��������
	BBS.Execute("update [Board] set orders=Orders + "&BoardNum&"+1  where RootID="&P_rs("RootID")&" And orders>"&P_rs("orders")&"")
	'���µ�ǰ��������
	If P_rs("parentstr")="0" Then
	BBS.Execute("update [Board] set Depth="&P_Rs("Depth")&"+1,orders="&P_Rs("orders")&"+1,rootid="&P_rs("Rootid")&",ParentID="&NewParentID&",ParentStr='" & P_Rs("BoardID") & "' Where BoardID="&BoardID)
	Else
	BBS.Execute("update [Board] set Depth="&P_Rs("Depth")&"+1,orders="&P_Rs("orders")&"+1,rootid="&P_rs("Rootid")&",ParentID="&NewParentID&",ParentStr='" & P_Rs("ParentStr") & ","& P_Rs("BoardID") &"' Where BoardID="&BoardID)
	End If
	Dim TempParentStr
	i=1
	'����������ͬʱ����ƶ�����i
	'����������������������
	'���Ϊԭ����ȼ��ϵ�ǰ������̳�����
	Set Rs=BBS.Execute("select * from [Board] where ParentStr like '%"&ParentStr & BoardID&"%' order by orders")
	Do while not rs.eof
	i=i+1
	If P_rs("parentstr")="0" Then'����丸��Ϊ�࣬��ô�������İ�������
		TempParentStr=P_rs("boardid") & "," & Replace(rs("parentstr"),ParentStr,"")
	Else
		TempParentStr=P_rs("parentstr") & "," & P_rs("boardID") & "," & replace(Rs("Parentstr"),ParentStr,"")
	End If
	BBS.Execute("update [Board] set depth=depth+"&P_rs("depth")&"-"&depth&"+1,orders="&P_rs("orders")&"+"&I&",Rootid="&P_rs("Rootid")&",ParentStr='"&TempParentStr&"' where BoardID="&Rs("BoardID"))
	Rs.movenext
	Loop
	Rs.Close
	Dim TempParentID,II
	TempParentID=NewParentID
	If RootID=P_rs("RootID") then'��ͬһ�������ƶ�
		'������ָ����ϼ���̳��������iΪ�����ƶ������İ�����
		'�����丸�������
		BBS.Execute("update [Board] set Child=child+"&i&" where (not ParentID=0) and BoardID="&TempParentID)
		For II=1 to P_Rs("depth")
			'�õ��丸��ĸ���İ���ID
			Set Rs=BBS.Execute("Select ParentID from [Board] where (not ParentID=0) and BoardID="&TempParentID)
			If Not (rs.eof and rs.bof) then
				TempParentid=Rs(0)
				'�����丸��ĸ��������
			BBS.Execute("update [Board] set Child=child+"&i&" where (not ParentID=0) and BoardID="&TempParentID)
			End if
		Next
		'������ԭ���������
		BBS.Execute("update [Board] set Child=child-"&i&" where (not ParentID=0) and BoardID="&ParentID)
		'������ԭ��������̳����
		For II=1 to Depth
			'�õ���ԭ����ĸ���İ���ID
			Set rs=BBS.Execute("select ParentID from [Board] where (not ParentID=0) and BoardID="&ParentID)
			if not (rs.eof and rs.bof) then
				ParentID=rs(0)
				'������ԭ����ĸ��������
				BBS.Execute("update [Board] set child=child-"&i&" where (not ParentID=0) and  boardid="&ParentID)
			End IF
		Next
	Else
	'������ָ����ϼ���̳��������iΪ�����ƶ������İ�����
	'�����丸�������
		BBS.Execute("update [Board] set Child=child+"&i&" where BoardID="&TempParentID)
		For II=1 to P_Rs("depth")
			'�õ��丸��ĸ���İ���ID
			Set Rs=BBS.Execute("Select ParentID from [Board] where BoardID="&TempParentID)
			If Not (rs.eof and rs.bof) then
				TempParentid=Rs(0)
				'�����丸��ĸ��������
			BBS.Execute("update [Board] set Child=child+"&i&" where  BoardID="&TempParentID)
			End if
		Next
	'������ԭ���������
	BBS.Execute("update [Board] set Child=child-"&i&" where BoardID="&ParentID)
	'������ԭ�����������������
	BBS.Execute("Update [Board] Set Orders=Orders-"&i&" where RootID="&RootID&" and Orders>"&Orders)
	'������ԭ��������̳����
		For II=1 to Depth
			'�õ���ԭ����ĸ���İ���ID
			Set rs=BBS.Execute("select ParentID from [Board] where BoardID="&ParentID)
			if not (rs.eof and rs.bof) then
				ParentID=rs(0)
				'������ԭ����ĸ��������
				BBS.Execute("update [Board] set child=child-"&i&" where boardid="&ParentID)
			End IF
		Next
	End if
	P_rs.Close:Set P_rs=Nothing
  End If
	Suc"","��̳�޸ĳɹ� !","Admin_Board.asp"
	BBS.NetLog"������̨_��̳�޸ĳɹ�!"
	BBS.Cache.clean("BoardInfo")
End Sub


Sub ClassOrders
	Dim BoardID
	Set Rs=BBS.Execute("Select BoardID,BoardName,RootID from[Board] where Depth=0 order by RootID")
	If Rs.Eof Then
		GoBack"","��̳û�з��࣡����<a href='?Action=AddClass'> ��ӷ���</a>"
		Exit Sub
	End If
	Response.Write"<form method=POST style='margin:0' action='?Action=SaveClassOrders'><div class='mian'><div class='top'>�������� </div><div class='divth'>����ʽ����С������������������д���������ֲ�����ͬ��</div>"
	Do while not rs.eof
	DIVTR Rs(1),"","<input name='BoardID' type='hidden' value='"&Rs(0)&"'><input name='RootID' type='hidden' value='"&Rs(2)&"'><input type=text name='NewRootID' value='"&Rs(2)&"' size='4' />",25,1
	Rs.MoveNext
	Loop
	Response.Write"<div class='bottom'><input type='submit' value='�� ��' class='button'><input type='reset' value='�� ��' class='button'></div></div></form>"
	Rs.Close
End Sub

Sub SaveClassOrders
	Dim BoardID,RootID,NewRootID,Temp,I
	Temp=","
	For i=1 to request.form("BoardID").count
		BoardID = request.form("BoardID")(i)
		RootID = request.form("RootID")(i)
		NewRootID = request.form("NewRootID")(i)
		If InStr(Temp,","&NewRootID&",")>0 Then 
			GoBack "�������","��������������ֲ���һ��!"
			Exit Sub
		End If
		Temp=Temp&NewRootID&","
		IF Not IsNumeric(BoardID) or Not isnumeric(NewRootID) Then
			GoBack "�������","����������д!"
			Exit Sub
		End IF
	Next
	For i=1 to request.form("BoardID").count
		BoardID = request.form("BoardID")(i)
		RootID = request.form("RootID")(i)
		NewRootID = request.form("NewRootID")(i)
		If RootID<>NewRootID Then
			BBS.Execute("Update [Board]Set RootID="&NewRootID&" where BoardID="&BoardID)
			Temp=BoardID
			BBS.Execute("Update [Board] Set RootID="&NewRootID&" where ParentStr like '%"&Temp&"%' And RootID="&RootID&"")
		End If
	Next
	Suc"","��������ɹ���","?"
	BBS.NetLog"������̨_��������ɹ���"	
	BBS.Cache.clean("BoardInfo")
End Sub


Sub BoardUnite
	Response.Write"<form method=POST style='margin:0' action='?Action=SaveBoardUnite'><div class='mian'><div class='top'>��̳�ϲ�</div>"&_
	"<div class='divth' style='padding:2px;'>����̳�� <select size='1' name='BoardID'><option value=''>��ѡ��ԭ��̳</option>"&BBS.BoardIDList(0,0)&"</select> �ϲ�����̳�� <select size='1' name='NewBoardID'><option value=''>��ѡ��Ŀ����̳</option>"&BBS.BoardIDList(0,0)&"</select> �� <input type='button' class='button' onclick=""if(confirm('�����󽫲��ָܻ�����ȷ��Ҫ�ϲ���'))form.submit()"" value='��̳�ϲ�'></div>"&_
	"<div class='divtr1' style='padding:5px;'><b>ע�����</b><font color=red>�˲������ɻָ��������ز�����</font><br>���಻�ܲ��������ܺ������µ���̳�ϲ���<br>�ϲ���ԭ��̳(����������̳)����ɾ������������(����������̳������)��ת�Ƶ�ָ����Ŀ����̳�� </div></div></form>"
End Sub

Sub SaveBoardUnite
	Dim BoardID,NewBoardID,TempParentStr,TempParentID,Rs1,S
	Dim I,AllTable
	Dim ParentStr,Depth,ParentID,Child,RootID
	BoardID=BBS.Fun.Getstr("BoardID")
	NewBoardID=BBS.Fun.Getstr("NewBoardID")
	IF BoardID="" Or NewBoardID="" Then
		GoBack"","����ָ����̳���ٽ��кϲ���"
		Exit Sub
	ElseIf BoardID=NewBoardID Then
		Goback"","ͬһ����̳���úϲ��ˣ�"
		Exit sub
	End If

	Set Rs=BBS.Execute("Select ParentStr,BoardID,Depth,ParentID,Child,RootID from [board] where BoardID="&BoardID)
	If Rs(2)="0" then
		Goback"ϵͳ����","���಻�����ϲ�������"
		Exit Sub
	End If
	ParentStr=Rs(0) & "," & Rs(1)
	ParentID=Rs(3)
	TempParentStr=rs(1)
	DepTh=rs(2)
	Child=rs(4)+1
	RootID=rs(5)
	Rs.Close
	TempParentID=ParentID
	'�ж��Ƿ�ϲ���������̳
	Set Rs=BBS.Execute("Select BoardID From [Board] where BoardID="&NewBoardID&" And ParentStr like '%"&ParentStr&"%'")
	If Not (rs.eof and rs.bof) then
		Goback"","���ܽ���̳�ϲ�����������̳��!"
		Exit Sub
	End if
	Rs.Close
	'�õ�ȫ��������̳ID
	i=0
	Set Rs=BBS.Execute("Select BoardID from [Board] where RootID="&RootID&" And ParentStr like '%"&ParentStr&"%'")
	do while not rs.eof
		If i=0 then
			TempParentStr=Rs(0)
		Else
			TempParentStr=TempParentStr & "," & Rs(0)
		End if
		i=i+1
		Rs.movenext
	loop
	If i>0 then
		ParentStr=TempParentStr & "," & BoardID
	Else
		ParentStr=BoardID
	End if
	'������ԭ��������̳������
	BBS.Execute("update [Board] set Child=Child-"&child&" where BoardID="&TempParentID)
	'������ԭ��������̳���ݣ������൱�ڼ�֦�����迼��
	For I=1 to Depth
		'�õ��丸��ĸ���İ���ID
		Set rs=BBS.Execute("select ParentID from [Board] where boardID="&TempParentID)
		If Not (rs.eof and rs.bof) then
			TempParentID=rs(0)
			BBS.Execute("update [Board] set Child=Child-"&Child&" where boardid="&TempParentID)
		End if
	Next
	'������̳��������
	AllTable=Split(BBS.BBStable(0),",")
	For i=0 To uBound(AllTable)
		BBS.Execute("update [BBS"&AllTable(i)&"] set BoardID="&NewBoardID&" where BoardID in ("&ParentStr&")")
	Next
	BBS.Execute("update [Topic] set BoardID="&NewBoardID&" where BoardID in ("&ParentStr&")")
	'ɾ�����ϲ���̳
	Set Rs=BBS.Execute("Select Sum(EssayNum),sum(TopicNum),sum(TodayNum) from [Board] where RootID="&RootID&" And BoardID in ("&ParentStr&")")
	BBS.Execute("Delete from [Board] where RootID="&RootID&" And BoardID in ("&ParentStr&")")
	'��������̳���Ӽ���
	BBS.Execute("update [Board] set EssayNum=EssayNum+"&rs(0)&",TopicNum=TopicNum+"&rs(1)&",TodayNum=TodayNum+"&rs(2)&" where BoardID ="&NewBoardID&"")
	'�����ϼ����
	set Rs1=BBS.Execute("select Depth,ParentStr,Boardid from [Board] where BoardID="&NewBoardID)
	If Rs1(0)>1 Then
	ParentStr=Rs1(0)
	BBS.Execute("update [Board] set EssayNum=EssayNum+"&rs(0)&",TopicNum=TopicNum+"&rs(1)&",TodayNum=TodayNum+"&rs(2)&" where boardid in ("&ParentStr&")")
	End If
	Rs1.Close:Set Rs1=Nothing
	Rs.Close
	'��������
	Set Rs=BBS.Execute("select name from [Admin] where BoardID="&NewBoardID&" or BoardID="&BoardID)
	If Not Rs.eof then
	do while not rs.eof
		S=S&Rs(0)&"|"
	Rs.movenext
	Loop
	S=left(S,len(S)-1)
	BBS.execute("update [Board] Set BoardAdmin='"&S&"' where BoardID="&NewBoardID)
	BBS.Execute("update [Admin] set boardID="&NewBoardID&" where boardID="&BoardID)
	End if
	Rs.Close
	Suc"","�ϲ��ɹ����Ѿ���ԭ��̳���������£����������Ӻϲ���Ŀ����̳��","?"
	BBS.NetLog"������̨_�ϲ���̳�ɹ�!"
	'���°�黺��
	BBS.Cache.clean("BoardInfo")
	BBS.Cache.clean("Board"&NewBoardID)
	BBS.Cache.clean("Board"&BoardID)
End Sub


Sub ClearData
	Set Rs=BBS.execute("Select BoardName,TopicNum,EssayNum From[Board] where BoardID="&BBS.BoardID&"")
	IF Rs.Eof Then
		GoBack"","��̳���治���ڣ����ܾ���ɾ����"
		Exit Sub
	End If
	Response.Write"<div class='mian'><div class='top'>"&Rs("BoardName")&" ��������</div>"&_
	"<div class='divtr1' style='padding:5px;'><b>�����������Ϣ</b><br />��������<span style='color:#F00'>"&Rs("TopicNum")&"</span>&nbsp;&nbsp;��������<span style='color:#F00'>"&Rs("EssayNum")&"</span>&nbsp;&nbsp;������������<span style='color:#F00'>"&BBS.Execute("Select Count(TopicID) From[Topic] where IsGood=1 and BoardID="&BBS.BoardID&"")(0)&"</span></div>"&_
	"<div class='divth' style='padding:3px'><form method=POST style='margin:0' action='?Action=StartClearData'>��� <b>"&Rs(0)&"</b> �� <input name='BoardID' value='"&BBS.BoardID&"' type='hidden' /><select name='SqlTableID'><option value='0'>�������ݱ�</option>"&SqlTableList&"</select> �� <input type='text' class='text' name='ClearDate' value='365' size='5'> ��ǰ�����ӡ� <input type='submit' value='ִ������' class='button'></form></div>"&_
	"<div class='divtr2' style='padding:5px;'><b>ע������</b><br><font color=red>�˲������ɻָ����������Ӳ��ᱻɾ����</Font><br>���������̳�����ڶִ࣬�д˲��������Ĵ����ķ�������Դ��<br>ִ�й��������ĵȺ����ѡ��ҹ���������ٵ�ʱ����¡�</div></div>"
	Rs.Close
End Sub

Sub StartClearData
	Dim SqlTableID,ClearDate,BoardID,AllTable,i,Temp
	SqlTableID=request.form("SqlTableID")
	ClearDate=request.form("ClearDate")
	BoardID=request.form("BoardID")
	If Not isnumeric(ClearDate) or Not isNumeric(SqlTableID) or Not isNumeric(BoardID)  Then
		GoBack"","����������д��"
		Exit Sub
	End If	
	IF Int(SqlTableID)=0 Then
		AllTable=Split(BBS.BBStable(0),",")
	Else
		AllTable=Split(SqlTableID,",")
	End if
	For i=0 to uBound(AllTable)
		Set Rs=BBS.Execute("Select TopicID,isVote From[Topic] where BoardID="&BoardID&" And IsGood=False And  DATEDIFF('d',[LastTime],'"&BBS.NowBbsTime&"')>"&ClearDate&" ")
		Do While Not Rs.Eof
		BBS.Execute("Delete from [Bbs"&AllTable(i)&"] where BoardID="&BoardID&" And (TopicId="&RS(0)&" Or ReplyTopicId="&RS(0)&")")
		IF Rs(1)=1 Then'ɾ��ͶƱ
			BBS.Execute("Delete from [TopicVote] where TopicID="&RS(0)&"")
			BBS.Execute("Delete from [TopicVoteUser] where TopicID="&RS(0)&"")
		End If
		Rs.movenext
		Loop
		Rs.Close
		BBS.Execute("Delete From[Topic] where BoardID="&BoardID&" And SqlTableID="&AllTable(i)&" And IsGood=0 And  DATEDIFF('d',[LastTime],'"&BBS.NowBbsTime&"')>"&ClearDate&" ")
	Next
	Temp=BBS.Execute("Select BoardName From[Board] where BoardID="&BoardID)(0)
	BBS.NetLog"������̨_������̳���棺"&Temp&" ��"&ClearDate&"��ǰ�����ݳɹ�!"
	Suc"","�ɹ�����������̳���ݣ�������һ��<a href='?Action=BoardUpdate'>��̳����</a>","?"
End Sub


Sub PassUser
	Dim Temp
	Set Rs=BBS.execute("Select PassUser,BoardName,Strings From [Board] where BoardID="&BBS.BoardID&" And ParentID<>0")
	IF Rs.eof Then
		GoBack"","����̳�����Ͳ�����֤��̳������������֤�û���"
		Exit Sub
	End If
	Temp=Split(Rs(2),"|")
	If Temp(6)="0" Then GoBack"","����̳�����Ͳ�����֤��̳������������֤�û���":Exit Sub
	Response.Write"<form method=POST style='margin:0' action='?Action=SavePassUser'><div class='mian'><div class='top'>�޸���̳��֤�û�</div>"
	DIVTR"������̳��","","<input name='BoardID' value='"&BBS.BoardID&"' type='hidden'>"&Rs("BoardName"),25,1
	DIVTR"ͨ����֤���û���","���û�֮���á�|������<br />�벻Ҫʹ�ûس���Enter","<textarea name='PassUser' rows='5'>"&Rs("PassUser")&"</textarea>",70,2
	Response.Write"<div class='bottom'><input type='submit' value='�� ��' class='button' /><input type='reset' class='button' value='�� ��' /></div></div></form>"
	Rs.Close
End Sub

Sub SavePassUser
	Dim PassUsers,BoardID
	BoardID=BBS.Fun.GetStr("BoardID")
	PassUsers=Trim(Replace(Request.Form("PassUser"),"'",""))
	PassUsers=Replace(PassUsers,chr(10), "")
	PassUsers=Replace(PassUsers,chr(13), "")
	BBS.Execute("Update [Board] Set PassUser='"&PassUsers&"' where BoardID="&BoardID&" And ParentID<>0")
	BBS.NetLog"������̨_������֤��Ա�ɹ�!"
	Suc"","�ɹ��ĸ����˸���̳����֤��Ա��","?"
	BBS.Cache.clean("BoardInfo")
End Sub


Sub BoardUpdate
	Response.Write"<div class='mian'>"&_
	"<div class='divth' style='height:50px'><b><div id='BBST'>���ݰ��������������Ե�</div></b><div style='margin:2px auto 0;width:400px;height:16px;background:#DEFAF1;text-align:left'><img src='Images/icon/hr1.gif' width=0 height='16' id='BBSimg' align='absmiddle' alt='������' /></div>"&_
	"<div><span id='BBStxt' style='font-size:9pt'>0</span>%</div></div></div>"
	Response.Flush
	Dim BoardNum,EssayNum,TopicNum,TodayNum,BoardAdmin,ParentStr,LastReply,LastCaption
	Dim AllTable,I,II,III,SQL,ReRs
	BoardNum=BBS.Execute("Select Count(BoardID) from[Board] Where ParentID<>0")(0)
	II=0
	Set Rs=BBS.Execute("Select BoardID,BoardName,Child,ParentStr,RootID from[Board] Where ParentID<>0  Order by Child,RootID,Orders Desc")
	If Not Rs.EOF Then 
	SQL=Rs.GetRows()
	Rs.Close
	For i=0 to UBound(SQL,2)
	EssayNum=0
	TopicNum=0
	TodayNum=0
	BoardAdmin=""
	LastReply="|||0||||||"
	LastCaption="��"
	AllTable=Split(BBS.BBStable(0),",")
	For III=0 To uBound(AllTable)
		EssayNum=EssayNum+BBS.Execute("Select Count(*) From[Bbs"&AllTable(III)&"] where BoardID="&SQL(0,i)&" And IsDel=0")(0)
		TodayNum=TodayNum+BBS.Execute("Select Count(*) From[Bbs"&AllTable(III)&"] where BoardID="&SQL(0,i)&" And IsDel=0 And DATEDIFF('d',[LastTime],'"&BBS.NowBbsTime&"')<1")(0)
	Next
	TopicNum=BBS.Execute("Select Count(TopicID) From[Topic] where BoardID="&SQL(0,i)&" and IsDel=0")(0)
	Set Rs=BBS.Execute("Select Name From[Admin] Where BoardID="&SQL(0,i)&"")
	Do While Not Rs.Eof
		BoardAdmin=BoardAdmin&Rs(0)&"|"
		Rs.Movenext
	Loop
	If BoardAdmin<>"" Then BoardAdmin=left(BoardAdmin,len(BoardAdmin)-1)
	Rs.Close
	Set Rs=BBS.execute("Select top 1 TopicID,Name,Caption,AddTime,Face,SqlTableID,ReplyNum From [Topic] where IsDel=0 And BoardID="&SQL(0,i)&" order by LastTime desc,TopicID desc")
	If Not Rs.eof then
	  If Rs(6) > 0 Then
	    Set ReRs=BBS.execute("Select top 1 TopicID,Name,Caption,AddTime,Face From [bbs"&Rs(5)&"] where ReplyTopicID="&Rs(0)&" And BoardID="&SQL(0,i)&" And IsDel=0 order by LastTime desc,BbsID desc")
	    If Not ReRs.eof then
		  LastCaption=replace(BBS.Fun.StrLeft(ReRs("Caption"),22),"'","''")
		  LastReply=ReRs("Name")&"|"&LastCaption&"|"&ReRs("AddTime")&"|"&ReRs("Face")&"|"&Rs("TopicID")&"|"&SQL(0,i)&"|"&Rs("SqlTableID")&""
		End If
		ReRs.close
	  Else
		LastCaption=replace(BBS.Fun.StrLeft(Rs("Caption"),22),"'","''")
		LastReply=Rs("Name")&"|"&LastCaption&"|"&Rs("AddTime")&"|"&Rs("Face")&"|"&Rs("TopicID")&"|"&SQL(0,i)&"|"&Rs("SqlTableID")&""
	  End If
	End If
	Rs.Close
	BBS.Execute("update [Board] Set EssayNum="&EssayNum&",TodayNum="&TodayNum&",TopicNum="&TopicNum&",BoardAdmin='"&BoardAdmin&"',LastReply='"&LastReply&"' where BoardID="&SQL(0,i)&"")
	'������ϼ���̳����ô�����ϼ���̳
	If SQL(2,I)>0 Then
		ParentStr=SQL(3,i) & "," & SQL(0,i)
		Set Rs=BBS.Execute("Select Sum(EssayNum),Sum(TopicNum),Sum(TodayNum) From [Board] Where ParentStr = '"&ParentStr&"'")
		If Not IsNull(Rs(0)) Then EssayNum = Rs(0) + EssayNum
		If Not IsNull(Rs(1)) Then TopicNum = Rs(1) + TopicNum
		If Not IsNull(Rs(2)) Then TodayNum = Rs(2) + TodayNum
		Rs.Close
		Set Rs=BBS.execute("Select top 1 TopicID,Name,Caption,AddTime,Face,SqlTableID,BoardID,ReplyNum From [Topic] where IsDel=0 And BoardID In ("&ParentStr&") Order by LastTime desc,LastTime Desc")
		If Not Rs.eof then
	      If Rs(7) > 0 Then
			Set ReRs=BBS.execute("Select top 1 TopicID,Name,Caption,AddTime,Face From [bbs"&Rs(5)&"] where ReplyTopicID="&Rs(0)&" And BoardID="&Rs("BoardID")&" And IsDel=0 order by LastTime desc,BbsID desc")
			If Not ReRs.eof then
		      LastCaption=replace(BBS.Fun.StrLeft(ReRs("Caption"),22),"'","''")
			  LastReply=ReRs("Name")&"|"&LastCaption&"|"&ReRs("AddTime")&"|"&ReRs("Face")&"|"&Rs("TopicID")&"|"&Rs("BoardID")&"|"&Rs("SqlTableID")&""
			End If
			ReRs.close
	      Else
			LastCaption=replace(BBS.Fun.StrLeft(Rs("Caption"),22),"'","''")
			LastReply=Rs("Name")&"|"&LastCaption&"|"&Rs("AddTime")&"|"&Rs("Face")&"|"&Rs("TopicID")&"|"&Rs("BoardID")&"|"&Rs("SqlTableID")&""
	      End If
		End If
		Rs.Close
		BBS.Execute("update [Board] Set EssayNum="&EssayNum&",TodayNum="&TodayNum&",TopicNum="&TopicNum&",LastReply='"&LastReply&"' where BoardID="&SQL(0,i)&"")
	End IF 
	If BoardAdmin="" Then
		BoardAdmin="��"
	Else
		BoardAdmin=Replace(Boardadmin,"|","��")
	End If
'���»���
	BBS.Cache.clean("Board"&SQL(0,i))
	Table "��̳ <Font color=blue>"&SQL(1,i)&"</Font> ����ɹ�","������"&EssayNum&" | ��������"&TopicNum&" | ��������"&TodayNum&" | ������"&BoardAdmin&" | �������⣺"&LastCaption&""
	II=II+1
	Response.Write "<script>document.getElementById(""BBSimg"").style.width=" & Fix((ii/BoardNum) * 400) & ";" & VbCrLf
	Response.Write "document.getElementById(""BBStxt"").innerHTML=""" & FormatNumber(ii/BoardNum*100,4,-1) & """;" & VbCrLf
	Response.Write "</script>" & VbCrLf
	Response.Flush
	Next
	End If
	Response.Write "<script>document.getElementById(""BBSimg"").style.width=400;document.getElementById(""BBStxt"").innerHTML=""100"";document.getElementById(""BBST"").innerHTML=""<font color=red>�ɹ��������</font>"";</script>"
	BBS.NetLog"������̨_������̳!"
End Sub

Function SqlTableList()
	Dim AllTable,I
	AllTable=Split(BBS.BBStable(0),",")
	For i=0 To uBound(AllTable)
	SqlTableList=SqlTableList&"<option value='"&AllTable(I)&"'>���ݱ�"&AllTable(I)&"</option>"
	Next
End Function


Sub Table(Str1,Str2)
	Response.Write"<div class='mian'><div class='divtr1' style='padding:5px;'>"&Str1&"<br>"&Str2&"</div></div>"
	Response.Flush
End Sub

Sub OrdersTopBoard
	Dim BoardID,ParentID,RootID,Orders,ParentStr,I,BoardNum,P_Rs
	BoardID=BBS.BoardID
	Set Rs=BBS.execute("Select Orders,ParentID,ParentStr From[Board] where BoardID="&BBS.BoardID)
	IF Rs.Eof or Rs.Bof Then
		GoBack"ϵͳ����","�ð��治���ڣ������Ѿ�ɾ���ˣ�"
		Exit Sub
	End If
	Orders=Rs(0)
	ParentID=Rs(1)
	ParentStr=Rs(2)
	Rs.Close
	'������Ϊ��ʱ
	If ParentID=0 then GoBack"ϵͳ����","����ID����":Exit Sub
	'�õ�������������
	ParentStr=ParentStr & ","
	BoardNum=BBS.Execute("select count(*) from [Board] where ParentStr like '%"&ParentStr & BoardID&"%'")(0)
	If Isnull(BoardNum) Then BoardNum=1
	'��ø�����Ϣ
	Set P_rs=BBS.Execute("select * from [board] where Boardid="&ParentID)
	'�ڻ���ƶ������İ����������������ָ����̳֮�����̳��������
	BBS.Execute("update [Board] set orders=Orders + "&BoardNum&"+1  where RootID="&P_rs("RootID")&" And orders>"&P_rs("orders")&"")
	'���µ�ǰ��������
	BBS.Execute("update [Board] set orders="&P_Rs("orders")&"+1 Where BoardID="&BoardID)
	Dim TempParentStr
	i=1
	'����������ͬʱ����ƶ�����i
	'����������������������
	Set Rs=BBS.Execute("select * from [Board] where ParentStr like '%"&ParentStr & BoardID&"%' order by orders")
	Do while not rs.eof
	i=i+1
	If P_rs("parentstr")="0" Then'����丸��Ϊ�࣬��ô�������İ�������
		TempParentStr=P_rs("boardid") & "," & Replace(rs("parentstr"),ParentStr,"")
	Else
		TempParentStr=P_rs("parentstr") & "," & P_rs("boardID") & "," & replace(Rs("Parentstr"),ParentStr,"")
	End If
	BBS.Execute("update [Board] set orders="&P_rs("orders")&"+"&I&",ParentStr='"&TempParentStr&"' where BoardID="&Rs("BoardID"))
	Rs.movenext
	Loop
	Rs.Close
	P_Rs.Close
	Set P_Rs=Nothing
	BBS.Cache.clean("BoardInfo")
	BBS.NetLog"������̨_������̳����ɹ�!"
	Response.Redirect"?"
	Response.End
End Sub
%>