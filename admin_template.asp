<!--#include file="Admin_Check.asp"-->
<script language="JavaScript">
<!--
var isopen=23;
function opendiv(i){
  if (isopen==i){
  document.getElementById("div"+i).style.display='none';
  isopen=23
  }else{
  document.getElementById("div"+i).style.display='block'
  document.getElementById("div"+isopen).style.display='none'
  isopen=i
  }
}

  // ѡ��ɫ
function SelectColor(what){
if(!document.all){alert("��ɫ�༭�������ã���ֱ����д��ɫ���뼴�ɡ�")}
else{
	var dEL = document.all("P"+what);
	var sEL = document.all("C"+what);
	var arr = showModalDialog("pic/edit/selcolor.htm", "", "dialogWidth:18em; dialogHeight:19em; status:0;help:0;scroll:no;");
	if (arr) {
		dEL.value=arr;
		sEL.style.backgroundColor=arr;
	}
	}
}
//-->
</script>
<%
Dim SkinsFlag,SkinsPIC
Dim Action,SkinConn,ID
CheckString "43"
ID=Request("ID")
SkinsFlag=Split("ҳ������|ҳ��ͷ��|���λ��|�ο���Ϣ|�û���Ϣ|���������|��ʾ���|�������Ϣ|�����������|��ʾ�����|��Ա����|��̳����|��ҳ����ͳ��|��ʾ�����б�|�����б���|��ʾ�������|��ʾ�����б�|���ӱ��|��ʾͶƱ|��ʾ����|�û��������|ͨ�����ݱ��|��̳����ͼ��|��������ͼ��|ҳ��ײ�","|")
SkinsPIC =Split("<font color=#A92D12>������ɫ������ɫ</font>|" &_
			"<font color=#A92D12>������ɫ������ɫ(1)</font>|"&_
			"<font color=#A92D12>������ɫ������ɫ(2)</font>|"&_
			
			"<font color=#513315>���״̬����ͨ��̳</font>|"&_
			"<font color=#513315>���״̬��������̳</font>|"&_
			"<font color=#513315>���״̬��������̳</font>|"&_
			"<font color=#513315>���״̬��������̳</font>|"&_
			
			"<font color=#04329B>������ť����������</font>|"&_
			"<font color=#04329B>������ť������ͶƱ</font>|"&_	
			"<font color=#04329B>������ť������ظ�</font>|"&_	
			
			"����״̬�����ö�|"&_
			"����״̬�����ö�|"&_
			"����״̬���ö�|"&_	
			"����״̬����������|"&_	
			"����״̬��ͶƱ����|"&_	
			"����״̬��������|"&_
			"����״̬�����ŵ�����|"&_
			"����״̬������������|"&_
			"����״̬��3Сʱ������|"&_
			
			"<font color=#836F38>�û�״̬������</font>|"&_	
			"<font color=#836F38>�û�״̬������</font>|"&_	
			
			"�����б�վ��|"&_	
			"�����б��ܰ���|"&_
			"�����б�����|"&_	
			"�����б�VIP��Ա|"&_	
			"�����б���Ա|"&_	
			"�����б������Ա|"&_	
			"�����б��ο�","|")	
Head()
Response.Write"<div class='mian'><div class='top'>��̳�������</div><div class='divth'>��<a href='Admin_Template.asp'>����б�</a>�� ��<a href='?Action=Add'>��ӷ��</a>�� ��<a href='?Action=Load'>������ݵ���</a>����<a href='?Action=SkinData'>������ݵ���</a>��</div></div>"
Select Case Request("Action")
Case"Add"
	Add
Case"SaveAdd"
	SaveAdd
Case"Del"
	Del
Case"Auto"
	Auto
Case"IsMode"
	IsMode
Case"Pass"
	pass
Case"Edit"
	Edit(0)
Case"UpdateName"
	UpdateName
Case"SaveEdit"
	SaveEdit
Case"EditPic"
	EditPic
Case"SkinData"
	SkinData
Case"Load"
	Load
Case"DataPost"
	DataPost
Case Else
	Main
End Select
Footer()

Sub Main
	Dim RsT,MainID,i
	With Response
	Set RsT=BBS.Execute("Select SkinID,SkinName,IsDefault,Ismode,Pass,remark From [Skins] Order By SkinID Asc")
	If RsT.Eof Then Exit Sub
	Rs=Rst.GetRows()
	RsT.CLose
	Set RsT=Nothing
	.write"<div class='mian'><div class='top'>����б�</div><table class='Stable' border='0' cellpadding='3' cellspacing='0'><tr><th width='25px'>ID</th><th width='18%'>�������</th><th>������</th></tr>"
	For i=0 To UBound(Rs,2)
		.write"<tr><td>"&Rs(0,i)&"</td><td title='"&Rs(5,i)&"'>"&Rs(1,i)&"</td><td>"
		If Rs(4,i)=1 Then
			.write "<A"
			If Rs(2,i)=1 Then .write " onClick=""alert('�˷����̳����ʹ���У���̳Ĭ�Ϸ���ܽ�ֹǰ̨��ʾ��');return false;"" "
			.write" HREF='?Action=Pass&ID="&Rs(0,i)&"'><FONT COLOR=red>�� ��ʾ</FONT></A>"
		Else
			.write " <A HREF='?Action=Pass&ID="&Rs(0,i)&"'>�� ��ʾ</A>"
		End If
		If Rs(2,i)=1 Then 
			.write " <FONT COLOR=red>�� ��̳Ĭ��</FONT>"
		Else
			.write " <A HREF='?Action=Auto&ID="&Rs(0,i)&"'>�� ��̳Ĭ��</A> "
		End IF
		If Rs(3,i)=1 Then
			.write " <A HREF='?Action=IsMode&ID="&Rs(0,i)&"'><FONT COLOR='red'>�� ����</FONT></A>"
		Else
			.write " <A HREF='?Action=IsMode&ID="&Rs(0,i)&"'>�� ����</A>"
		End IF
		.write" <A HREF='?Action=EditPic&ID="&Rs(0,i)&"'>"&IconE&" ��̬ͼƬ</A>"
		.write" <A HREF='?Action=Edit&ID="&Rs(0,i)&"'>"&IconE&" ҳ��ṹ</A>"
		.write" <A HREF='#this' onClick="""
	IF Rs(2,i)=1 then
			.write "alert('�˷����̳����ʹ���У���̳Ĭ�Ϸ����ɾ����')"
		else
			.write "checkclick('ɾ���󽫲��ָܻ�����ȷ��Ҫɾ����','?Action=Del&ID="&Rs(0,i)&"')"
		end if
		.write """>"&IconD&" ɾ��</a></td></tr>"
	Next
	.write"</table></div>"
	End With
End Sub

Sub UpdateName
	If Request("SkinName")="" Then Goback"","":Exit Sub
	BBS.Execute("Update [Skins] Set SkinName='"&Replace(Request("SkinName"),"'","")&"',Remark='"&Replace(Request("Remark"),"'","")&"' Where SkinID="&ID)
	Suc"","��������޸ĳɹ���","?"
	BBS.NetLog"������̨_�޸ķ������"
End Sub

Sub Add
	Dim Temp
	Set RS=BBS.Execute("Select Top 1 SkinName From [Skins] Where IsMode=1")
	If Not Rs.Eof Then
		Temp="��ǰ���� <span color=red>"&Rs("SkinName")&"</span> �ķ���ͼƬ��ģ��ṹ"
	Else
		Temp="��ǰû������ ���ģ�� "
	End If
	Rs.Close
	Response.Write"<FORM METHOD=POST style='margin:0' ACTION='?Action=SaveAdd'><div class='mian'><div class='top'>����·�� </div><div class='divth'>"&Temp&"</div>"
	DIVTR"������ƣ�","","<INPUT NAME='SkinName' TYPE='text' class='text' size='12' maxlength='50'>",25,1
	DIVTR"���Ŀ¼��","","<INPUT NAME='SkinDir' TYPE='text' class='text' size='12' maxlength='50'><br />���ñ����ͼƬ��Ŀ¼����Ŀ¼Ϊ \Skins\Default ��ֻ��д Default����д�󽫲����޸�",25,1
	DIVTR"���ע��","","<INPUT NAME='Remark' type='text' class='text' size='60' maxlength='255' >",50,1
	Response.Write"<div class='bottom'><input type='submit' class='button' value='ȷ����������һ��'></div></div></form>"
End Sub

Sub SaveAdd
	Dim Temp,Content,PIC,Txt,i,SkinDir
	If Request("SkinDir")="" or Request("SkinName")="" or Request("Remark")="" Then GoBack"","":Exit Sub
	Set RS=BBS.Execute("Select Top 1 SkinName,Content,PIC From [Skins] Where IsMode=1")
	If Not Rs.Eof Then
		Content=Rs(1)
		PIC=Rs(2)
		Txt="��ǰ���� <font color='#F00'>"& Rs(0) &"</font> �ķ���ͼƬ��ģ��ṹ"
	Else
		For i = 0 to Ubound(Skinsflag)
			Content=Content&VBCrlf&"["&Skinsflag(i)&"]"&VBCrlf&"[/"&Skinsflag(i)&"]"&VBCrlf
		Next
		PIC="|||||||||||||||||||||||||||||||||"
		Txt="��ǰû�����÷��ģ��,����ĸ��Ϊ�ա�"
	End If
	Rs.Close
	BBS.Execute("Insert Into [Skins](SkinName,Remark,Content,Pic,SkinDir,isDefault,ismode,Pass) values('"&Replace(Request("SkinName"),"'","''")&"','"&Replace(Left(Request("Remark"),255),"'","''")&"','"&Replace(Content,"'","''")&"','"&PIC&"','"&Request("SkinDir")&"',0,0,1)")
	Showtable "������һ��","�ɹ���� <b>"&Request("SkinName")&"</b> ���<br />���ڱ༭���Ľṹ-->><br />"&txt
	BBS.NetLog"������̨_��ӷ��"
	ID=Conn.Execute("Select Max(SkinID) from [Skins]")(0)
	Edit(1)
End Sub

Sub Edit(flag)
	Dim Temp,SkinName,HelpTxt,I,flagname,Remark
	Set RS=BBS.Execute("Select SkinName,Content,Remark From [Skins] Where SkinID="&ID)
	SkinName=Rs(0)
	BBS.Skins=Rs(1)
	Remark = Rs(2)
	Rs.Close
	Temp="<FORM METHOD=POST style='margin:0 ' ACTION='?Action=SaveEdit'>"
	If Flag=1 Then
	  Response.Write Temp&"<input name='Add' type='hidden' value='1' />"
	Else
	  Response.Write"<div class='mian'><div class='top'>��������Ϣ</div>"
	  Response.Write"<div class='divth' style='height:25px'><FORM METHOD='POST' ACTION='?Action=UpdateName'><B>������ƣ�</B><INPUT TYPE='text' class='text' NAME='SkinName' value='"&SkinName&"' maxlength='50'> <INPUT TYPE='hidden' name='ID' value='"&ID&"'></div>"
	  Response.Write"<div class='divth' style='height:25px'><B>���ע��</B><INPUT TYPE='text' class='text' NAME='Remark' value='"&Remark&"' size='60' maxlength='255'></div>"
	  Response.Write"<div class='divth' style='height:25px'><INPUT TYPE='submit' value='���ķ������' class='button'></FORM></div>"&Temp
	End If
	Response.Write"<div class='mian'><div class='top'>���ҳ��ṹ</div>"
	Response.Write"<INPUT TYPE='hidden' name='ID' value='"&ID&"'>"
	For i = 0 to Ubound(Skinsflag)
	If Skinsflag(i)="��ʾ���" or Skinsflag(i)="�������Ϣ" or Skinsflag(i)="��ʾ�����" or Skinsflag(i)="��ʾ�����б�" or Skinsflag(i)="��ʾ����" Then
	FlagName="<font color=#5C481D>&nbsp;(ѭ��)</font>"
	Else
	FlagName=""
	End If
		Temp=BBS.Readskins(Skinsflag(I))
		Response.Write"<div onMouseOver=this.style.backgroundColor='#FFFFFF' onMouseOut=this.style.backgroundColor='' class='divtr1' style='line-height:24px'> <div style='float:right;width:50%;'><a href=#this onClick='javascript:opendiv("&i&")'>"&IconE&"�༭����</a></div><div style='color:#F00'>["&Skinsflag(i)&"]"&FlagName&"</div></div>"
		Response.Write"<div class='divth' id='div"&i&"' style='height:213px;color:#999999;display:none'><div style=' float:left;width:18px'><br /><b>"&Skinsflag(i)&"</b></div><div style='margin-left:18px;'><TEXTAREA NAME='TmpName_"&i&"' ROWS='16'  style='width:100%'>"&BBS.Readskins(Skinsflag(i))&"</TEXTAREA></div></div>"
	Next
	Response.Write"<a id='div23'></a><div class='bottom'><input class='button' type='submit' value=' ȷ���ύ '><input class='button' type='reset' value=' ȡ����д '></div></form></div>"
End Sub

Sub SaveEdit()
	Dim Temp,Content,ResultErr,i
	For i = 0 to Ubound(Skinsflag)
		Content=Content&"["&Skinsflag(i)&"]"&Request("TmpName_"&i)&"[/"&Skinsflag(i)&"]"
		If Request("TmpName_"&i)="" Then ResultErr=ResultErr&"<FONT COLOR=#FF0033>["&Skinsflag(i)&"]</FONT><br />"
	Next
	BBS.Execute("update [Skins] set Content='"&Replace(Content,"'","''")&"' where SkinID="&ID&"")
	If Request.Form("Add")="1" Then
		Showtable"������һ��","�ɹ�������ģ��ṹ�����ڱ༭���ͼƬ-->>"
		EditPic()
	Else
		If ResultErr<>"" Then
			Suc"","�ɹ�������ģ��,�������µ�Ԫ�أ�<br />"&ResultErr&" ��û������!<li>�뵽����������༭��</li>","?"
		Else
			Suc"","�ɹ�������ģ��","?"
		End If
	End If
	BBS.Cache.Clean("Skin_"& ID)
	BBS.NetLog"������̨_�޸ķ�����"
End Sub

Sub EditPic()
	Dim Temp,Pic,i,SkinName
	IF Request("PIC6")="" Then
		Set RS=BBS.Execute("Select SkinName,Pic From [Skins] Where SkinID="&ID)
		If not Rs.Eof Then
			SkinName=Rs(0)
			pic=Rs(1)
		Rs.Close
		Else
			Goback"","�Ҳ���������¼�����ݣ������Ѿ�ɾ��"
			Exit Sub
		End If
		If Pic<>"" Then
		Pic=Split(PIC,"|")
		Else
		Pic=Split("|||||||||||||||||||||||||||","|")
		End if
		Response.Write"<FORM METHOD=POST style='margin:0' ACTION='?Action=EditPic&ID="&ID&"'><div class='mian'><div class='top'>"&SkinName&" &nbsp;&nbsp;&nbsp�༭��ɫ/ͼƬ</div><div class='divtr2' style='height:38px;padding:5px'>˵������ɫ��Ҫ����ģ������ĵط�������Ϊ͸��ɫ������ģ����ƥ�䡣<br />ͼƬ���������ͼƬ����Ҳ���������ִ��档<br />ͼƬ����������<u>&lt;img src=&quot;Skins/20051201/user.gif&quot;  border=&quot;0&quot;&gt;</u></div>"
		DIVTR SkinsPIC(0),"","<input name='PIC0' type='text' class='text' size='8' value='"&Replace(PIC(0),"'","")&"' /><img border='0' align=""absmiddle"" src='pic/edit/rect.gif' width='18' name='CIC0' style='cursor:pointer;background-color:"&Replace(PIC(0),"'","")&"'  onClick=""SelectColor('IC0')""> <span class='explain'>������ģ�������������ɫ</span>",22,1
		DIVTR SkinsPIC(1),"","<input name='PIC1' type='text' class='text' size='8' value='"&Replace(PIC(1),"'","")&"' /><img border='0' align=""absmiddle"" src='pic/edit/rect.gif' width='18' name='CIC1' style='cursor:pointer;background-color:"&Replace(PIC(1),"'","")&"'  onClick=""SelectColor('IC1')""> <span class='explain'>������ģ������ı�����ǳɫ</span>",22,2
		DIVTR SkinsPIC(2),"","<input name='PIC2' type='text' class='text' size='8' value='"&Replace(PIC(2),"'","")&"' /><img border='0' align=""absmiddle"" src='pic/edit/rect.gif' width='18' name='CIC2' style='cursor:pointer;background-color:"&Replace(PIC(2),"'","")&"'  onClick=""SelectColor('IC2')""> <span class='explain'>ͬ�ϣ���������ĵ���ɫ����һЩ</span>",22,2
		For i = 3 to Ubound(SkinsPIC)
		DIVTR SkinsPIC(i),"","<input name='PIC"&i&"' type='text' class='text'  size='55' style='width:98%' value='"&Replace(PIC(i),"'","&#39")&"' />",22,1
		next
		Response.Write"<div class='bottom'><input type='submit' value=' �� �� ' class='button'></div></div></FORM>"
	ELse
		For i = 0 to Ubound(SkinsPIC)
		PIC=PIC&Replace(Request.Form("PIC"&i),"|","&#124")&"|"
		Next
		BBS.Execute("Update [Skins] Set PIC='"&Replace(PIC,"'","''")&"' Where SkinID="&ID)
		Suc"","����ͼƬ�޸ĳɹ���","?"
		BBS.Cache.Clean("Skin_"& ID)
		BBS.NetLog"������̨_�޸ķ��ͼƬ"
	End If
End Sub

Sub Auto
	Dim Temp
	BBS.Execute("Update [Config] Set SkinID="&ID)
	BBS.Execute("Update [Skins] Set IsDefault=0")
	BBS.Execute("Update [Skins] Set IsDefault=1 where SkinID="&ID )
	'���»���
	If BBS.Cache.Valid("parameter") Then
		Temp=Split(BBS.Cache.Value("parameter"),"<$$>")
		BBS.Cache.Add "parameter",Replace(Join(Temp,"<$$>"),"<$$>"&Temp(2)&"<$$>","<$$>"&ID&"<$$>"),dateadd("n",2000,BBS.NowBBSTime)
	End If
	Suc"","�����Ϊ��̳Ĭ��ʹ�óɹ���","?"
End Sub

Sub IsMode
	If BBS.Execute("Select IsMode From [Skins] where SkinID="&ID)(0)=0 Then 
		BBS.Execute("Update [Skins] Set IsMode=0")
		BBS.Execute("Update [Skins] Set IsMode=1 where SkinID="&ID )
		Suc"","�˷������Ϊ�����̳��������ģ�棡","?"
	Else
		BBS.Execute("Update [Skins] Set IsMode=0 where SkinID="&ID )
		Suc"","�Ѿ��ɹ�ȡ������Ϊ�����̳��������ģ�棡","?"
	End If
End Sub

Sub Pass
Dim s
	If BBS.Execute("Select Pass From [Skins] where SkinID="&ID)(0)=0 Then 
		BBS.Execute("Update [Skins] Set Pass=1 where SkinID="&ID )
		Suc"","�ɹ��Ŀ����˷��,�� <a href='Admin_Confirm.asp?action=setjsmenu'>�ؽ�ǰ̨�˵�</a> ","?"
		BBS.NetLog"������̨_���������ʾ��"
	Else
		BBS.Execute("Update [Skins] Set Pass=0 where SkinID="&ID )
		Suc"","�ɹ��Ľ�ֹ�˸÷����ǰ̨����ʾ���� <a href='Admin_Confirm.asp?action=setjsmenu'>�ؽ�ǰ̨�˵�</a>","?"
		BBS.NetLog"������̨_������ò���ʾ"
	End IF
End Sub

Sub Del
	BBS.Execute("Delete From [Skins] Where SkinID="&ID)
	BBS.Cache.clean("Skin_"& ID)
	Suc"","����ѱ��ɹ�ɾ����","?"
	BBS.NetLog"������̨_ɾ�����"
End Sub

Sub Load()
	Response.Write"<form action='?action=SkinData&Flag=Load' method='post'><div class='mian'><div class='top'>������ģ������</div>"
	DIVTR"������ģ�����ݿ�����","","<input name='skinmdb' type='text' class='text' size='30' value='Skins/Skins.mdb'>",25,1
	Response.Write"<div class='bottom'><input type='submit' class='button' value='��һ��' /></div></div></form>"
End Sub

Sub DataPost
	Dim Msg,MdbName,S,Temp
	IF ID="" Then GoBack"","����û��ѡ��һ����Ŀ��":Exit Sub
	MdbName=request("SkinMdb")
	SkinConnection(mdbname)
    If Request("To")="InputSkin" Then
	    If Request.Form("DelFlag")="1" Then
	       SkinConn.Execute("Delete * From [Skins] Where SkinID In ("&ID&")")
		   Suc "","�ɹ��İ�"&mdbname&"�ķ��ģ��ɹ�ɾ����","?":Exit Sub
		Else
		  Set Rs=SkinConn.Execute("select SkinName,Content,Pic,remark,SkinDir from [Skins] where SkinID in ("&ID&")  order by SkinID ")
          While Not Rs.Eof
			  Temp=Replace(Rs(0),"'","''")
			  If Not BBS.Execute("Select * From [Skins] where SkinName='"&Temp&"'").Eof Then Temp=Temp&"(��)"
              BBS.Execute("Insert Into [Skins](SkinName,Content,Pic,Remark,SkinDir,isdefault,ismode,Pass) values('"&Temp&"','"&Replace(Rs(1),"'","''")&"','"&Replace(Rs(2),"'","''")&"','"&Replace(Rs(3),"'","''")&"','"&Replace(Rs(4),"'","''")&"',0,0,0)")  
			  Rs.Movenext
          Wend
		  	Rs.Close
		  S="���ģ�����ݵ���ɹ���"
		End If
    Else
	      Set Rs=BBS.Execute(" select SkinName,Content,Pic,remark,SkinDir from [Skins] where SkinID in ("&ID&")  order by SkinID ")
          While Not Rs.Eof
              SkinConn.Execute("Insert Into [Skins](SkinName,Content,Pic,remark,SkinDir) values('"&Replace(Rs(0),"'","''")&"','"&Replace(Rs(1),"'","''")&"','"&Replace(Rs(2),"'","''")&"','"&Replace(Rs(3),"'","''")&"','"&Replace(Rs(4),"'","''")&"')") 
			  Rs.Movenext
          Wend 
		  Rs.Close
		  S="���ģ�����ݵ����ɹ���"
   End If
	SkinConn.Close
	Set SkinConn=Nothing
   	BBS.NetLog"������̨_"&S
	Suc"",S,"?"
End Sub


Sub SkinData
	Dim Title,FlagName,MdbName,act
	If Request("Flag")="Load" Then
		FlagName="����"
		act="InputSkin"
		MdbName=trim(Request.form("SkinMdb"))
		Title="������ģ������ ��"&MdbName&"���ݿ��еķ���б�"
		If MdbName="" Then
			GoBack"","����д������ģ��ķ��ר�����ݿ⣡"
			Exit Sub
		End If
	Else
		FlagName="����"
		act="OutSkin"
		Title="������̳���еķ��ģ������"
	End If
	If act="InputSkin" Then
		SkinConnection(MdbName)
		On error resume next
		Set Rs=SkinConn.Execute("select SkinID,SkinName,Content,Pic,remark,SkinDir from [Skins] order by SkinID")
		if err Then
		err.Clear
		GoBack"","�˷�����ݿ�İ汾�뵱ǰ�İ汾�����ݣ�":Exit Sub
		End If
	Else
		Set Rs=BBS.Execute("select SkinID,SkinName,Content,Pic,remark,SkinDir from [Skins] order by SkinID")
		MdbName="Skins/Skins.mdb"
	End If
	Dim Temp,i
	IF Rs.Eof Then
		GoBack"","�����ݿ���û�з��ģ������ݣ�":Exit Sub
	End IF
	Temp=Rs.GetRows()
	Response.Write"<form action='Admin_Template.asp?Action=DataPost&To="&Act&"' method='post'><div class='mian'><div class='top'>"&Title&"</div>"
	Response.Write"<div class='divth'><div class='divtd1' style='width:35px'><b>ѡ��</b></div><div class='divtd2' style='width:20%'>�������</div><div class='divtd2'>��Ϣ����</div><div style='clear: both;'></div></div>"
	For i=0 To Ubound(Temp,2)
		Response.Write"<div class='divtr1' style='overflow:hidden; height:25px'><div class='divtd1' style='width:35px'><input type='checkbox' name='ID' value='"&Temp(0,i)&"' /></div><div class='divtd2' style='width:20%'>"&Temp(1,i)&"</div><div class='divtd2'>"&Temp(4,i)&"</div><div style='clear: both;'></div></div>"
	Next
	Response.Write"<div class='bottom'>"&FlagName&"�����ݿ⣺<input type='text' class='text' name='SkinMdb' size='30' value='"&MdbName&"' /> <input type='submit' class='button' value='"&FlagName&"' />"
	If act="InputSkin" Then
		Response.Write"<input name='DelFlag' type='hidden' value='0' /><input type='button' class='button' value=ɾ��  onClick=""if(confirm('ɾ���󽫲��ָܻ�����ȷ��Ҫɾ����')){form.DelFlag.value=1;form.submit()}"" />"
	End If
	Response.Write"<input type='checkbox' name='chkall' value='on' onClick='CheckAll(this.form)'>ȫѡ</div></div></form>"
End Sub

Sub SkinConnection(Mdbname)
	On Error Resume Next 
	Set SkinConn = Server.CreateObject("ADODB.Connection")
	SkinConn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(MdbName)
	If Err Then 
		GoBack"",Mdbname&" ���ݿⲻ���ڣ���ȷ�����·���Ƿ���ȷ�����û�з����ʱ���ݿ⣬�뵽BBS�ٷ�</a>����"
		Footer()
		Response.end
	End If
End Sub
%>