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

  // 选颜色
function SelectColor(what){
if(!document.all){alert("颜色编辑器不可用，请直接填写颜色代码即可。")}
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
SkinsFlag=Split("页面属性|页面头部|你的位置|游客信息|用户信息|版块分区表格|显示版块|最后发帖信息|版块分区简洁表格|显示简洁版块|会员生日|论坛联盟|首页在线统计|显示在线列表|主题列表表格|显示版块在线|显示主题列表|帖子表格|显示投票|显示帖子|用户控制面版|通用内容表格|论坛属性图标|帖子属性图标|页面底部","|")
SkinsPIC =Split("<font color=#A92D12>定义颜色：线条色</font>|" &_
			"<font color=#A92D12>定义颜色：表面色(1)</font>|"&_
			"<font color=#A92D12>定义颜色：表面色(2)</font>|"&_
			
			"<font color=#513315>版块状态：普通论坛</font>|"&_
			"<font color=#513315>版块状态：限制论坛</font>|"&_
			"<font color=#513315>版块状态：特殊论坛</font>|"&_
			"<font color=#513315>版块状态：锁定论坛</font>|"&_
			
			"<font color=#04329B>发帖按钮：发表帖子</font>|"&_
			"<font color=#04329B>发帖按钮：发表投票</font>|"&_	
			"<font color=#04329B>发帖按钮：发表回复</font>|"&_	
			
			"帖子状态：总置顶|"&_
			"帖子状态：区置顶|"&_
			"帖子状态：置顶|"&_	
			"帖子状态：精华主题|"&_	
			"帖子状态：投票主题|"&_	
			"帖子状态：热门帖|"&_
			"帖子状态：开放的主题|"&_
			"帖子状态：锁定的主题|"&_
			"帖子状态：3小时内新帖|"&_
			
			"<font color=#836F38>用户状态：在线</font>|"&_	
			"<font color=#836F38>用户状态：离线</font>|"&_	
			
			"在线列表：站长|"&_	
			"在线列表：总版主|"&_
			"在线列表：版主|"&_	
			"在线列表：VIP会员|"&_	
			"在线列表：会员|"&_	
			"在线列表：隐身会员|"&_	
			"在线列表：游客","|")	
Head()
Response.Write"<div class='mian'><div class='top'>论坛风格设置</div><div class='divth'>【<a href='Admin_Template.asp'>风格列表</a>】 【<a href='?Action=Add'>添加风格</a>】 【<a href='?Action=Load'>风格数据导入</a>】【<a href='?Action=SkinData'>风格数据导出</a>】</div></div>"
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
	.write"<div class='mian'><div class='top'>风格列表</div><table class='Stable' border='0' cellpadding='3' cellspacing='0'><tr><th width='25px'>ID</th><th width='18%'>风格名称</th><th>风格管理</th></tr>"
	For i=0 To UBound(Rs,2)
		.write"<tr><td>"&Rs(0,i)&"</td><td title='"&Rs(5,i)&"'>"&Rs(1,i)&"</td><td>"
		If Rs(4,i)=1 Then
			.write "<A"
			If Rs(2,i)=1 Then .write " onClick=""alert('此风格论坛正在使用中，论坛默认风格不能禁止前台显示！');return false;"" "
			.write" HREF='?Action=Pass&ID="&Rs(0,i)&"'><FONT COLOR=red>√ 显示</FONT></A>"
		Else
			.write " <A HREF='?Action=Pass&ID="&Rs(0,i)&"'>× 显示</A>"
		End If
		If Rs(2,i)=1 Then 
			.write " <FONT COLOR=red>√ 论坛默认</FONT>"
		Else
			.write " <A HREF='?Action=Auto&ID="&Rs(0,i)&"'>× 论坛默认</A> "
		End IF
		If Rs(3,i)=1 Then
			.write " <A HREF='?Action=IsMode&ID="&Rs(0,i)&"'><FONT COLOR='red'>√ 引用</FONT></A>"
		Else
			.write " <A HREF='?Action=IsMode&ID="&Rs(0,i)&"'>× 引用</A>"
		End IF
		.write" <A HREF='?Action=EditPic&ID="&Rs(0,i)&"'>"&IconE&" 动态图片</A>"
		.write" <A HREF='?Action=Edit&ID="&Rs(0,i)&"'>"&IconE&" 页面结构</A>"
		.write" <A HREF='#this' onClick="""
	IF Rs(2,i)=1 then
			.write "alert('此风格论坛正在使用中，论坛默认风格不能删除！')"
		else
			.write "checkclick('删除后将不能恢复！您确定要删除吗？','?Action=Del&ID="&Rs(0,i)&"')"
		end if
		.write """>"&IconD&" 删除</a></td></tr>"
	Next
	.write"</table></div>"
	End With
End Sub

Sub UpdateName
	If Request("SkinName")="" Then Goback"","":Exit Sub
	BBS.Execute("Update [Skins] Set SkinName='"&Replace(Request("SkinName"),"'","")&"',Remark='"&Replace(Request("Remark"),"'","")&"' Where SkinID="&ID)
	Suc"","风格名称修改成功。","?"
	BBS.NetLog"操作后台_修改风格名称"
End Sub

Sub Add
	Dim Temp
	Set RS=BBS.Execute("Select Top 1 SkinName From [Skins] Where IsMode=1")
	If Not Rs.Eof Then
		Temp="当前引用 <span color=red>"&Rs("SkinName")&"</span> 的风格的图片和模版结构"
	Else
		Temp="当前没有引用 风格模版 "
	End If
	Rs.Close
	Response.Write"<FORM METHOD=POST style='margin:0' ACTION='?Action=SaveAdd'><div class='mian'><div class='top'>添加新风格 </div><div class='divth'>"&Temp&"</div>"
	DIVTR"风格名称：","","<INPUT NAME='SkinName' TYPE='text' class='text' size='12' maxlength='50'>",25,1
	DIVTR"风格目录：","","<INPUT NAME='SkinDir' TYPE='text' class='text' size='12' maxlength='50'><br />放置本风格图片的目录，如目录为 \Skins\Default 则只填写 Default，填写后将不能修改",25,1
	DIVTR"风格备注：","","<INPUT NAME='Remark' type='text' class='text' size='60' maxlength='255' >",50,1
	Response.Write"<div class='bottom'><input type='submit' class='button' value='确定！进入下一步'></div></div></form>"
End Sub

Sub SaveAdd
	Dim Temp,Content,PIC,Txt,i,SkinDir
	If Request("SkinDir")="" or Request("SkinName")="" or Request("Remark")="" Then GoBack"","":Exit Sub
	Set RS=BBS.Execute("Select Top 1 SkinName,Content,PIC From [Skins] Where IsMode=1")
	If Not Rs.Eof Then
		Content=Rs(1)
		PIC=Rs(2)
		Txt="当前引用 <font color='#F00'>"& Rs(0) &"</font> 的风格的图片和模版结构"
	Else
		For i = 0 to Ubound(Skinsflag)
			Content=Content&VBCrlf&"["&Skinsflag(i)&"]"&VBCrlf&"[/"&Skinsflag(i)&"]"&VBCrlf
		Next
		PIC="|||||||||||||||||||||||||||||||||"
		Txt="当前没有引用风格模版,下面的各项都为空。"
	End If
	Rs.Close
	BBS.Execute("Insert Into [Skins](SkinName,Remark,Content,Pic,SkinDir,isDefault,ismode,Pass) values('"&Replace(Request("SkinName"),"'","''")&"','"&Replace(Left(Request("Remark"),255),"'","''")&"','"&Replace(Content,"'","''")&"','"&PIC&"','"&Request("SkinDir")&"',0,0,1)")
	Showtable "进入下一步","成功添加 <b>"&Request("SkinName")&"</b> 风格<br />现在编辑风格的结构-->><br />"&txt
	BBS.NetLog"操作后台_添加风格"
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
	  Response.Write"<div class='mian'><div class='top'>风格基本信息</div>"
	  Response.Write"<div class='divth' style='height:25px'><FORM METHOD='POST' ACTION='?Action=UpdateName'><B>风格名称：</B><INPUT TYPE='text' class='text' NAME='SkinName' value='"&SkinName&"' maxlength='50'> <INPUT TYPE='hidden' name='ID' value='"&ID&"'></div>"
	  Response.Write"<div class='divth' style='height:25px'><B>风格备注：</B><INPUT TYPE='text' class='text' NAME='Remark' value='"&Remark&"' size='60' maxlength='255'></div>"
	  Response.Write"<div class='divth' style='height:25px'><INPUT TYPE='submit' value='更改风格名称' class='button'></FORM></div>"&Temp
	End If
	Response.Write"<div class='mian'><div class='top'>风格页面结构</div>"
	Response.Write"<INPUT TYPE='hidden' name='ID' value='"&ID&"'>"
	For i = 0 to Ubound(Skinsflag)
	If Skinsflag(i)="显示版块" or Skinsflag(i)="最后发帖信息" or Skinsflag(i)="显示简洁版块" or Skinsflag(i)="显示主题列表" or Skinsflag(i)="显示帖子" Then
	FlagName="<font color=#5C481D>&nbsp;(循环)</font>"
	Else
	FlagName=""
	End If
		Temp=BBS.Readskins(Skinsflag(I))
		Response.Write"<div onMouseOver=this.style.backgroundColor='#FFFFFF' onMouseOut=this.style.backgroundColor='' class='divtr1' style='line-height:24px'> <div style='float:right;width:50%;'><a href=#this onClick='javascript:opendiv("&i&")'>"&IconE&"编辑内容</a></div><div style='color:#F00'>["&Skinsflag(i)&"]"&FlagName&"</div></div>"
		Response.Write"<div class='divth' id='div"&i&"' style='height:213px;color:#999999;display:none'><div style=' float:left;width:18px'><br /><b>"&Skinsflag(i)&"</b></div><div style='margin-left:18px;'><TEXTAREA NAME='TmpName_"&i&"' ROWS='16'  style='width:100%'>"&BBS.Readskins(Skinsflag(i))&"</TEXTAREA></div></div>"
	Next
	Response.Write"<a id='div23'></a><div class='bottom'><input class='button' type='submit' value=' 确定提交 '><input class='button' type='reset' value=' 取消重写 '></div></form></div>"
End Sub

Sub SaveEdit()
	Dim Temp,Content,ResultErr,i
	For i = 0 to Ubound(Skinsflag)
		Content=Content&"["&Skinsflag(i)&"]"&Request("TmpName_"&i)&"[/"&Skinsflag(i)&"]"
		If Request("TmpName_"&i)="" Then ResultErr=ResultErr&"<FONT COLOR=#FF0033>["&Skinsflag(i)&"]</FONT><br />"
	Next
	BBS.Execute("update [Skins] set Content='"&Replace(Content,"'","''")&"' where SkinID="&ID&"")
	If Request.Form("Add")="1" Then
		Showtable"进入下一步","成功更改了模版结构，现在编辑风格图片-->>"
		EditPic()
	Else
		If ResultErr<>"" Then
			Suc"","成功更改了模版,但是以下的元素：<br />"&ResultErr&" 还没有内容!<li>请到风格管理里面编辑！</li>","?"
		Else
			Suc"","成功更改了模版","?"
		End If
	End If
	BBS.Cache.Clean("Skin_"& ID)
	BBS.NetLog"操作后台_修改风格代码"
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
			Goback"","找不到这条记录的数据，可能已经删除"
			Exit Sub
		End If
		If Pic<>"" Then
		Pic=Split(PIC,"|")
		Else
		Pic=Split("|||||||||||||||||||||||||||","|")
		End if
		Response.Write"<FORM METHOD=POST style='margin:0' ACTION='?Action=EditPic&ID="&ID&"'><div class='mian'><div class='top'>"&SkinName&" &nbsp;&nbsp;&nbsp编辑颜色/图片</div><div class='divtr2' style='height:38px;padding:5px'>说明：颜色主要用于模板以外的地方，留空为透明色，请与模板相匹配。<br />图片各项除了用图片代码也可以用文字代替。<br />图片代码样例：<u>&lt;img src=&quot;Skins/20051201/user.gif&quot;  border=&quot;0&quot;&gt;</u></div>"
		DIVTR SkinsPIC(0),"","<input name='PIC0' type='text' class='text' size='8' value='"&Replace(PIC(0),"'","")&"' /><img border='0' align=""absmiddle"" src='pic/edit/rect.gif' width='18' name='CIC0' style='cursor:pointer;background-color:"&Replace(PIC(0),"'","")&"'  onClick=""SelectColor('IC0')""> <span class='explain'>用于在模版以外的线条颜色</span>",22,1
		DIVTR SkinsPIC(1),"","<input name='PIC1' type='text' class='text' size='8' value='"&Replace(PIC(1),"'","")&"' /><img border='0' align=""absmiddle"" src='pic/edit/rect.gif' width='18' name='CIC1' style='cursor:pointer;background-color:"&Replace(PIC(1),"'","")&"'  onClick=""SelectColor('IC1')""> <span class='explain'>用于与模版以外的表格表面浅色</span>",22,2
		DIVTR SkinsPIC(2),"","<input name='PIC2' type='text' class='text' size='8' value='"&Replace(PIC(2),"'","")&"' /><img border='0' align=""absmiddle"" src='pic/edit/rect.gif' width='18' name='CIC2' style='cursor:pointer;background-color:"&Replace(PIC(2),"'","")&"'  onClick=""SelectColor('IC2')""> <span class='explain'>同上，但比上面的的颜色略深一些</span>",22,2
		For i = 3 to Ubound(SkinsPIC)
		DIVTR SkinsPIC(i),"","<input name='PIC"&i&"' type='text' class='text'  size='55' style='width:98%' value='"&Replace(PIC(i),"'","&#39")&"' />",22,1
		next
		Response.Write"<div class='bottom'><input type='submit' value=' 提 交 ' class='button'></div></div></FORM>"
	ELse
		For i = 0 to Ubound(SkinsPIC)
		PIC=PIC&Replace(Request.Form("PIC"&i),"|","&#124")&"|"
		Next
		BBS.Execute("Update [Skins] Set PIC='"&Replace(PIC,"'","''")&"' Where SkinID="&ID)
		Suc"","风格的图片修改成功。","?"
		BBS.Cache.Clean("Skin_"& ID)
		BBS.NetLog"操作后台_修改风格图片"
	End If
End Sub

Sub Auto
	Dim Temp
	BBS.Execute("Update [Config] Set SkinID="&ID)
	BBS.Execute("Update [Skins] Set IsDefault=0")
	BBS.Execute("Update [Skins] Set IsDefault=1 where SkinID="&ID )
	'更新缓存
	If BBS.Cache.Valid("parameter") Then
		Temp=Split(BBS.Cache.Value("parameter"),"<$$>")
		BBS.Cache.Add "parameter",Replace(Join(Temp,"<$$>"),"<$$>"&Temp(2)&"<$$>","<$$>"&ID&"<$$>"),dateadd("n",2000,BBS.NowBBSTime)
	End If
	Suc"","风格设为论坛默认使用成功！","?"
End Sub

Sub IsMode
	If BBS.Execute("Select IsMode From [Skins] where SkinID="&ID)(0)=0 Then 
		BBS.Execute("Update [Skins] Set IsMode=0")
		BBS.Execute("Update [Skins] Set IsMode=1 where SkinID="&ID )
		Suc"","此风格被设置为添加论坛风格的引用模版！","?"
	Else
		BBS.Execute("Update [Skins] Set IsMode=0 where SkinID="&ID )
		Suc"","已经成功取消了作为添加论坛风格的引用模版！","?"
	End If
End Sub

Sub Pass
Dim s
	If BBS.Execute("Select Pass From [Skins] where SkinID="&ID)(0)=0 Then 
		BBS.Execute("Update [Skins] Set Pass=1 where SkinID="&ID )
		Suc"","成功的开启了风格,请 <a href='Admin_Confirm.asp?action=setjsmenu'>重建前台菜单</a> ","?"
		BBS.NetLog"操作后台_风格设置显示！"
	Else
		BBS.Execute("Update [Skins] Set Pass=0 where SkinID="&ID )
		Suc"","成功的禁止了该风格在前台的显示！请 <a href='Admin_Confirm.asp?action=setjsmenu'>重建前台菜单</a>","?"
		BBS.NetLog"操作后台_风格设置不显示"
	End IF
End Sub

Sub Del
	BBS.Execute("Delete From [Skins] Where SkinID="&ID)
	BBS.Cache.clean("Skin_"& ID)
	Suc"","风格已被成功删除！","?"
	BBS.NetLog"操作后台_删除风格"
End Sub

Sub Load()
	Response.Write"<form action='?action=SkinData&Flag=Load' method='post'><div class='mian'><div class='top'>导入风格模版数据</div>"
	DIVTR"导入风格模版数据库名：","","<input name='skinmdb' type='text' class='text' size='30' value='Skins/Skins.mdb'>",25,1
	Response.Write"<div class='bottom'><input type='submit' class='button' value='下一步' /></div></div></form>"
End Sub

Sub DataPost
	Dim Msg,MdbName,S,Temp
	IF ID="" Then GoBack"","您还没有选定一个项目！":Exit Sub
	MdbName=request("SkinMdb")
	SkinConnection(mdbname)
    If Request("To")="InputSkin" Then
	    If Request.Form("DelFlag")="1" Then
	       SkinConn.Execute("Delete * From [Skins] Where SkinID In ("&ID&")")
		   Suc "","成功的把"&mdbname&"的风格模版成功删除！","?":Exit Sub
		Else
		  Set Rs=SkinConn.Execute("select SkinName,Content,Pic,remark,SkinDir from [Skins] where SkinID in ("&ID&")  order by SkinID ")
          While Not Rs.Eof
			  Temp=Replace(Rs(0),"'","''")
			  If Not BBS.Execute("Select * From [Skins] where SkinName='"&Temp&"'").Eof Then Temp=Temp&"(新)"
              BBS.Execute("Insert Into [Skins](SkinName,Content,Pic,Remark,SkinDir,isdefault,ismode,Pass) values('"&Temp&"','"&Replace(Rs(1),"'","''")&"','"&Replace(Rs(2),"'","''")&"','"&Replace(Rs(3),"'","''")&"','"&Replace(Rs(4),"'","''")&"',0,0,0)")  
			  Rs.Movenext
          Wend
		  	Rs.Close
		  S="风格模版数据导入成功！"
		End If
    Else
	      Set Rs=BBS.Execute(" select SkinName,Content,Pic,remark,SkinDir from [Skins] where SkinID in ("&ID&")  order by SkinID ")
          While Not Rs.Eof
              SkinConn.Execute("Insert Into [Skins](SkinName,Content,Pic,remark,SkinDir) values('"&Replace(Rs(0),"'","''")&"','"&Replace(Rs(1),"'","''")&"','"&Replace(Rs(2),"'","''")&"','"&Replace(Rs(3),"'","''")&"','"&Replace(Rs(4),"'","''")&"')") 
			  Rs.Movenext
          Wend 
		  Rs.Close
		  S="风格模版数据导出成功！"
   End If
	SkinConn.Close
	Set SkinConn=Nothing
   	BBS.NetLog"操作后台_"&S
	Suc"",S,"?"
End Sub


Sub SkinData
	Dim Title,FlagName,MdbName,act
	If Request("Flag")="Load" Then
		FlagName="导入"
		act="InputSkin"
		MdbName=trim(Request.form("SkinMdb"))
		Title="导入风格模版数据 在"&MdbName&"数据库中的风格列表："
		If MdbName="" Then
			GoBack"","请填写导入风格模版的风格专用数据库！"
			Exit Sub
		End If
	Else
		FlagName="导出"
		act="OutSkin"
		Title="导出论坛现有的风格模版数据"
	End If
	If act="InputSkin" Then
		SkinConnection(MdbName)
		On error resume next
		Set Rs=SkinConn.Execute("select SkinID,SkinName,Content,Pic,remark,SkinDir from [Skins] order by SkinID")
		if err Then
		err.Clear
		GoBack"","此风格数据库的版本与当前的版本不兼容！":Exit Sub
		End If
	Else
		Set Rs=BBS.Execute("select SkinID,SkinName,Content,Pic,remark,SkinDir from [Skins] order by SkinID")
		MdbName="Skins/Skins.mdb"
	End If
	Dim Temp,i
	IF Rs.Eof Then
		GoBack"","该数据库中没有风格模版的数据！":Exit Sub
	End IF
	Temp=Rs.GetRows()
	Response.Write"<form action='Admin_Template.asp?Action=DataPost&To="&Act&"' method='post'><div class='mian'><div class='top'>"&Title&"</div>"
	Response.Write"<div class='divth'><div class='divtd1' style='width:35px'><b>选择</b></div><div class='divtd2' style='width:20%'>风格名称</div><div class='divtd2'>信息描述</div><div style='clear: both;'></div></div>"
	For i=0 To Ubound(Temp,2)
		Response.Write"<div class='divtr1' style='overflow:hidden; height:25px'><div class='divtd1' style='width:35px'><input type='checkbox' name='ID' value='"&Temp(0,i)&"' /></div><div class='divtd2' style='width:20%'>"&Temp(1,i)&"</div><div class='divtd2'>"&Temp(4,i)&"</div><div style='clear: both;'></div></div>"
	Next
	Response.Write"<div class='bottom'>"&FlagName&"的数据库：<input type='text' class='text' name='SkinMdb' size='30' value='"&MdbName&"' /> <input type='submit' class='button' value='"&FlagName&"' />"
	If act="InputSkin" Then
		Response.Write"<input name='DelFlag' type='hidden' value='0' /><input type='button' class='button' value=删除  onClick=""if(confirm('删除后将不能恢复！您确定要删除吗？')){form.DelFlag.value=1;form.submit()}"" />"
	End If
	Response.Write"<input type='checkbox' name='chkall' value='on' onClick='CheckAll(this.form)'>全选</div></div></form>"
End Sub

Sub SkinConnection(Mdbname)
	On Error Resume Next 
	Set SkinConn = Server.CreateObject("ADODB.Connection")
	SkinConn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(MdbName)
	If Err Then 
		GoBack"",Mdbname&" 数据库不存在！请确认你的路径是否正确，如果没有风格临时数据库，请到BBS官方</a>下载"
		Footer()
		Response.end
	End If
End Sub
%>