<!--#include file="Admin_Check.asp"-->
<%
Const MaxDepth=5 '可以自由设置最大论坛等级深度，默认为5级论坛
Head()
CheckString "11"
Response.Write "<div class='mian'><div class='top'>论坛版面管理</div><div class='divth'>【<a href='?'>论坛管理</a>】【<a href='?Action=AddClass'>增加分类</a>】【<a href='?Action=AddBoard'>增加论坛</a>】【<a href='?Action=ClassOrders'>分类排序</a>】【<a href='?Action=BoardUpdate'>论坛整理</a>】【<a href='?Action=BoardUnite'>论坛合并</a>】</div></div>"
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
	Response.Write"<div class='mian'><div class='top'>论坛版块</div><table class='Stable' border='0' cellpadding='3' cellspacing='0'><tr><th width='180px'>论坛版面</th><th width='70px'>类别</th><th>相应操作</th></tr>"
	Set Rs=BBS.execute("Select BoardID,BoardName,ParentID,Depth,Child,Strings from [board] order by Rootid,orders")
	If Rs.Bof Then
	    Response.Write "</table></div>"
		GoBack "","论坛没有分类！请先 <a href='Admin_Board.asp?Action=AddClass'>添加分类</a>"
		Exit Sub
	End If
	Brs=Rs.GetRows(-1)
	Rs.close
	For I=0 To Ubound(Brs,2)
		Temp="<tr><td>"
		Install="<a href='?Action=AddBoard&BoardID="&BRs(0,i)&"'>"&IconA&"添加论坛</a>"
		If Brs(3,i)=0 Then'分类
			BoardTypeName="分类"
			Temp="<tr><td>"
			If Brs(4,i)>0 Then'如果有子论坛
				Temp=Temp&Brs(1,i)&" ("&Brs(4,i)&")"
			Else
				Temp=Temp&Brs(1,i)
			End If
			Install=Install &"<a href='?Action=EditClass&BoardID="&Brs(0,i)&"'>"&IconE&"编辑分类</a> "
			If Brs(4,i)>0 Then
				Install=Install &"<a href=""javascript:alert('不能删除！该分类含有论坛!\n\n要删除本类，必须先把属下的论坛删除或移走。')"">"
			Else
				Install=Install &"<a href=""javascript:checkclick('删除后将不能恢复！您确定要删除吗？','?Action=DelClass&BoardID="&Brs(0,i)&"')"">"
			End If
			Install=Install&IconD&"删除分类</a>"
		Else'版面
			Strings=Split(Brs(5,i),"|")
		If Strings(7)="1" Then
			BoardTypeName="锁定论坛"
		ElseIf Strings(6)="1" or Strings(5)="1" Then
			BoardtypeName="特殊论坛"
		ElseIf Strings(9)="1" or Strings(3)="1" Then
			BoardTypeName="限制论坛"
		Else
			BoardtypeName="普通论坛"
		End If
		If Strings(0)="1" Then BoardtypeName=BoardtypeName&"(类)"
			Po=""
			For II=1 To Brs(3,i)
				Po=Po&"<font color=red>|</Font> "
			Next
			If Brs(4,i)>0 Then'如果有子论坛
				Temp=Temp&Po&Brs(1,i)&" ("&Brs(4,i)&")"
			Else
				Temp=Temp&Po&Brs(1,i)
			End If
			Install=Install &"<a href='?Action=EditBoard&BoardID="&Brs(0,i)&"'>"&IconE&"版面设置</a>"
			If Brs(4,i)>0 Then
				Install=Install &" <a href=""javascript:alert('不能删除！该版面含有子论坛!\n\n要删除本版，必须先把属下的子论坛删除或移走。')"">"&IconD&"删除版面</a>"
			Else
				Install=Install &" <a href=""javascript:checkclick('删除后将不能恢复！您确定要删除吗？','?Action=DelBoard&BoardID="&Brs(0,i)&"')"">"&IconD&"删除论坛</a>"
			End If
				Install=Install &" <a href='?Action=ClearData&BoardID="&BRs(0,i)&"'><img src='images/Icon/recycle.gif' border='0' align='absmiddle' /> 清理数据</a>"
				Install=Install & " <a href='?Action=OrdersTopBoard&BoardID="&BRs(0,i)&"'><img src='Images/icon/Top.gif' border='0' align='absmiddle' /> 排序置上</a>"
			If Strings(6)="1" Then
				Install=Install & " <a href='?Action=PassUser&BoardID="&BRs(0,i)&"'><img src='Images/icon/user.gif' border='0' align='absmiddle' /> 认证用户</a>"
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
	Response.Write"<form method=POST style=""margin:0"" action=""?Action=SaveClass""><div class='mian'><div class='top'>添加分类</div>"
	DIVTR"分类名称：","论坛的分类名称","<input name=""NewBoardID"" type=""hidden"" value='"&NewBoardID&"' /><input type=""text"" class='text' class='text' name='BoardName' size='30'>",40,1
	DIVTR"论坛简洁显示：","设置版面的下属论坛是否以简洁方式显示","<input name='s2' type='radio' value='0' checked>否<input name='s2' type='radio' value='1'>是",40,1
	DIVTR"简洁显示个数：","当设置为简洁显示，每一行显示的个数[一般显示4个比较美观]","<input name='s3' type='text' class='text' size='3' maxlength='2' value='4'>个",40,1
	Response.Write"<div class='bottom'><input type=""submit"" class='button' value=""提 交""><input type=""reset"" value=""重 置"" class='button'></div></div></form>"
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
			GoBack"内部系统出错","不能指定和别的论坛一样的序号，如果不能解决此问题，请到BBS官方论坛寻求帮助！"
			Exit Sub
		End if
		Rs.Close
		Set Rs=BBS.Execute("Select Max(RootID) From [Board]")
		MaxRootID=Rs(0)+1
		If isnull(MaxRootID) then MaxRootID=1
		Rs.Close
		BBS.execute("Insert into [Board](BoardName,BoardID,RootID,Depth,ParentID,Orders,Child,ParentStr,Strings)Values('"&BoardName&"',"&NewBoardID&","&MaxRootID&",0,0,0,0,'0','0|"&s1&"|"&s2&"|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0')")
		BBS.Cache.clean("BoardInfo")
		Temp="添加了论坛分类 <b>"&BoardName&" </b> 成功!"
		BBS.NetLog "操作后台_"&Temp
		Suc"",Temp,"?"
	End If
End Sub

Sub EditClass
	Dim BoardID,Rs,Strings
	Set Rs=BBS.Execute("Select BoardName,Strings from[Board] where BoardID="&BBS.BoardID&"")
	If Rs.Eof Then
		GoBack "系统出错","论坛找不到这个分类，可能已经删除了。":Exit Sub
	End If
	Strings=Split(Rs(1),"|")
	Response.Write"<form method=POST style=""margin:0"" action=""?Action=SaveEditClass""><div class='mian'><div class='top'>编辑分类</div>"
	DIVTR"分类名称：","修改论坛分类的名称","<input name=""BoardID"" type=""hidden"" value='"&BBS.BoardID&"' /><input type=""text"" class='text' class='text' name='BoardName' size='30' value='"&Rs(0)&"' />",40,1
	DIVTR"论坛简洁显示：","设置版面的下属论坛是否以简洁方式显示",GetRadio("s1","否",Strings(1),0)&GetRadio("s1","是",Strings(1),1),40,1
	DIVTR"简洁显示个数：","当设置为简洁显示，每一行显示的个数[一般显示4个比较美观]","<input name='s2' type='text' class='text' size='3' maxlength='2' value='"&Strings(2)&"'>个",40,1
	Response.Write"<div class='bottom'><input type=""submit"" value=""提 交"" class='button'><input type=""reset"" value=""重 置"" class='button'></div></div></form>"
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
			GoBack"系统出错","论坛找不到这个分类，可能已经删除了。":Exit Sub
		End if
		BBS.execute("Update [Board] Set BoardName='"&BoardName&"',Strings='0|"&s1&"|"&s2&"|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0' where BoardID="&BoardID)
		Temp="论坛分类名称改为 <b>"&BoardName&"</b> 成功!"
		BBS.NetLog "操作后台_"&Temp
		BBS.Cache.clean("BoardInfo")
		Suc "",Temp,"?"
	End If
End Sub

Sub DelClass
	Dim Temp
	IF BBS.Execute("Select Count(BoardID) From[Board] where ParentID="&BBS.BoardID)(0)=0 Then
		BBS.Execute("Delete From[Board] where ParentID=0 And BoardID="&BBS.BoardID)
		BBS.Cache.clean("BoardInfo")
		Temp="删除论坛分类成功!"
		BBS.NetLog "操作后台_"&Temp
		Suc"",Temp,"?"
	End If
End Sub

Sub DelBoard
	Dim AllTable,I,II,Depth,ParentID,RootID,Orders,Temp
	Set Rs=BBS.Execute("Select Depth,ParentID,RootID,Orders,Child From[Board] where BoardID="&BBS.BoardID)
	If Rs.Eof Then 
		Goback"","不存在，论坛可能已经删除了 !"
		Exit Sub
	ElseIf Rs(4)>0 Then
		Goback"","该论坛含有属下论坛，不能删除 !"
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
	'删除主题记录
	BBS.Execute("Delete From[Topic] where BoardID="&BBS.BoardID)
	BBS.Execute("Delete From[Board] where BoardID="&BBS.BoardID)
	'更新其父类的版面数
	BBS.Execute("update [Board] set Child=child-1 where BoardID="&ParentID)
	'更新其父类的其它版面排序
	BBS.Execute("Update [Board] Set Orders=Orders-1 where RootID="&RootID&" and Orders>"&Orders)
	'更新其所属论坛数据
		For II=1 to Depth
			'得到其父类的父类的版面ID
			Set rs=BBS.Execute("select ParentID from [Board] where BoardID="&ParentID)
			if not (rs.eof and rs.bof) then
				ParentID=rs(0)
				'更新其父类的父类版面数
				BBS.Execute("update [Board] set child=child-1 where boardid="&ParentID)
			End IF
			Rs.Close
		Next
	BBS.Cache.clean("BoardInfo")
	BBS.Cache.clean("Board"&BBS.BoardID)
	Temp="成功的删除论坛版面 (包括该论坛的所有帖子)!"
	BBS.NetLog "操作后台_"&Temp
	Suc"",Temp,"Admin_Board.asp"
End Sub


Sub AddBoard
	If BBS.execute("Select BoardID from [Board] where Depth=0").Eof Then
		GoBack"","没有分类不能添加论坛！请先 <a href=Admin_Board.asp?Action=AddClass>添加分类</a>"
		Exit Sub
	End if
	Response.Write"<form method=POST style='margin:0' action='?Action=SaveBoard'>"
	Response.Write"<div class='mian'><div class='top'>论坛添加 </div>"
	DIVTR"属于分类或论坛：","选择要属于那个分类或那个论坛","<select name='ParentID'>"&BBS.BoardIDList(BBS.BoardID,20)&"</select>*",40,1
	DIVTR"论坛名称：","论坛版面的名称","<input type='text' class='text' name='BoardName' size='30' />*",40,2
	DIVTR"标志图片：","论坛版面Logo地址，为了首页美观一定要填写","<input type='text' class='text' name='BoardImg' size='30' />*",40,1
	DIVTR"论坛介绍：","论坛版面描述","<textarea rows='3' name='Introduce'  cols='60'></textarea>*",58,2
	DIVTR"作为小类：","是否设置为类，设置后该版面不能发帖","<input name='s0' type='radio' value='0' checked>否<input name='s0' type='radio' value='1'>是",40,1
	DIVTR"论坛简洁显示：","设置版面的下属论坛是否以简洁方式显示","<input name='s1' type='radio' value='0' checked>否<input name='s1' type='radio' value='1'>是",40,2
	DIVTR"简洁显示个数：","当设置为简洁显示，每一行显示的个数[一般显示4个比较美观]","<input name='s2' type='text' class='text' size='3' value='4' maxlength='2'>个",40,1
	DIVTR"论坛类型：","设置论坛的类型，可以多选","<input type='checkbox' name='s3' value='1' />会员（只有会员才能浏览帖子） <br /><input type='checkbox' name='s4' value='1' />只读（可以浏览帖子，但只有站长、超版、版主能发帖）<br /><input type='checkbox' name='s5' value='1' />VIP（只有vip用户才能进入）<br /><input type='checkbox' name='s6' value='1' />认证（只有通过认证的用户才能进入）",90,2
	DIVTR"锁定论坛：","论坛除了站长外一律不得进入","<input name='s7' type='radio' value='0' checked>开放<input name='s7' type='radio' value='1'>锁定",40,1
	DIVTR"设置限制：","用户达到这些资源便可以进入","帖数：<input name='s10' type='text' class='text' value='0' size='6' /><br>"&BBS.Info(121)&"：<input name='s11' type='text' class='text' value='0' size='6' /><br>"&BBS.Info(120)&"：<input name='s12' type='text' class='text' value='0' size='6' /><br>"&BBS.Info(122)&"：<input name='s13' type='text' class='text' value='0' size='6' />",120,2
	DIVTR"上传设置：","当论坛系统默认禁止上传后，将不起作用。","<input type='radio' name='s14' value='0' />禁止 <input type='radio' name='s14' value='1' checked />全部会员 <input type='radio' name='s14' value='2' />只有VIP会员",40,1
	Response.Write"<div class='bottom'><input type='submit' value=' 提 交 ' class='button' /><input type='reset' value=' 重 置 ' class='button' /></div></div></form>"
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
		If Not Isnumeric(strings(i)) Then GoBack"","一些项目必需用数字填写！":Exit Sub
	Next
	'如果设置限制资源
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
		GoBack"","没有分类不能添加论坛！请先<a href='Admin_Board.asp?Action=AddClass'>添加分类</a>"
		Exit Sub
	Else
		NewBoardID=Rs(0)+1
	End If
	Rs.Close
	Set Rs=BBS.execute("Select RootID,Depth,Child,Orders,ParentStr,ParentID From[Board] where BoardID="&ParentID&"")
	IF Rs.Eof or Rs.Bof Then
		GoBack"系统程式出错！","没有指定父类或父论坛！"
		Exit Sub
	End If
	RootID=Rs(0)
	Depth=Rs(1)
	Child=Rs(2)
	Orders=Rs(3)
	ParentStr=Rs(4)
	Rs.Close
	If Depth+1>MaxDepth Then
		GoBack "","考滤到论坛的实用易用，本论坛限制了最多只能有" & MaxDepth & "级论坛。^_^"
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
		'当上级分类深度大于0的时候要更新其父类（或父类的父类）的版面数和相关排序
		For i=1 to Depth
			'更新其父类版面数
			BBS.Execute("update [Board] set Child=Child+1 where BoardID="&parentID)
			'得到其父类的父类的版面ID
			Set rs=BBS.Execute("select ParentID from [Board] where BoardID="&parentID)
			If not (rs.eof and rs.bof) then
				ParentID=rs(0)
			End if
			Rs.Close
			'当循环次数大于1并且运行到最后一次循环的时候直接进行更新
			If i=depth then
			BBS.Execute("update [Board] set Child=Child+1 where BoardID="&parentID)
			End if
		next
		'更新该版面排序以及大于本需要和同在本分类下的版面排序序号
		BBS.Execute("update [Board] set Orders=orders+1 where RootID="&RootID&" And orders>"&orders)
		BBS.Execute("update [Board] set Orders="&Orders&"+1 where BoardID="&NewBoardID&"")
	Else
		'当上级分类深度为0的时候只要更新上级分类版面数
		BBS.Execute("update [Board] set child=child+1 where Boardid="&ParentID)
		Set rs=BBS.Execute("select max(Orders) from [Board]")
		BBS.Execute("update [Board] set Orders="&rs(0)&"+1 where BoardID="&NewBoardID )
		Rs.Close
	End if
	End if
	If Strings(6)="1" Then
		Suc"","成功的添加了论坛 <b>"&BoardName&"</b> !<li>此论坛为认证论坛，暂时只有最高管理员能够进入。<li>你可以通过 <a href=?Action=PassUser>管理</a> 项目来添加可以进入该论坛的用户","Admin_Board.asp"	
	Else
		Suc"","成功的添加了论坛 <b>"&BoardName&"</b> !","Admin_Board.asp"	
	End IF
	BBS.NetLog"操作后台_添加论坛<b>"&BoardName&"</b>成功!"
	BBS.Cache.clean("BoardInfo")
	BBS.Cache.clean("Board"&NewBoardID)
End Sub

Sub EditBoard
	Dim BoardName,Strings,Introduce,BoardImg,ParentID
	Dim Temp,Chk
	Set Rs=BBS.execute("Select ParentID,BoardName,Strings,Introduce,BoardImg From[Board] Where BoardID="&BBS.BoardID&"")
	If Rs.eof Then
		GoBack"","该版面不存在，可能已经删除了"
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
	Response.Write"<div class='mian'><div class='top'>编辑论坛 </div><input name='BoardID' type='hidden' value='"&BBS.BoardID&"'>"
	
	DIVTR"论坛名称：","论坛版面的名称","<input type='text' class='text' name='BoardName' size='30' value='"&BoardName&"' />*",40,1
	DIVTR"属于分类或论坛：","选择要属于那个分类或那个论坛","<select name='ParentID'>"&BBS.BoardIDList(ParentID,20)&"</select>*",40,2
	DIVTR"标志图片：","论坛版面Logo地址，为了首页美观一定要填写","<input type='text' class='text' name='BoardImg' size='30' value='"&BoardImg&"' />*",40,1
	DIVTR"论坛介绍：","论坛版面描述","<textarea rows='3' name='Introduce'  cols='60'>"&Introduce&"</textarea>*",58,2
	DIVTR"作为小类：","是否设置为类，设置后该版面不能发帖",GetRadio("s0","否",Strings(0),0)&GetRadio("s0","是",Strings(0),1),40,1
	DIVTR"论坛简洁显示：","设置版面的下属论坛是否以简洁方式显示",GetRadio("s1","否",Strings(1),0)&GetRadio("s1","是",Strings(1),1),40,2
	DIVTR"简洁显示个数：","当设置为简洁显示，每一行显示的个数[一般显示4个比较美观]","<input name='s2' type='text' class='text' size='3' value='"&Strings(2)&"' maxlength='2'>个",40,1
	If Strings(3)="1" Then Chk="checked" Else Chk=""
	Temp="<input type='checkbox' name='s3' value='1' "&Chk&" />会员（只有会员才能浏览帖子） <br />"
	If Strings(4)="1" Then Chk="checked" Else Chk=""
	Temp=Temp&"<input type='checkbox' name='s4' value='1' "&Chk&" />只读（可以浏览帖子，但只有站长、超版、版主能发帖）<br />"
	If Strings(5)="1" Then Chk="checked" Else Chk=""
	Temp=Temp&"<input type='checkbox' name='s5' value='1' "&Chk&" />VIP（只有vip用户才能进入）<br />"
	If Strings(6)="1" Then Chk="checked" Else Chk=""
	Temp=Temp&"<input type='checkbox' name='s6' value='1' "&Chk&" />认证（只有通过认证的用户才能进入）"
	DIVTR"论坛类型：","设置论坛的类型，可以多选",Temp,90,2
	DIVTR"锁定论坛：","论坛除了站长外一律不得进入",GetRadio("s7","开放",Strings(7),0)&GetRadio("s7","锁定",Strings(7),1),40,1
	DIVTR"设置限制：","用户达到这些资源便可以进入","帖数：<input name='s10' type='text' class='text' value='"&Strings(10)&"' size='6' /><br>"&BBS.Info(121)&"：<input name='s11' type='text' class='text' value='"&Strings(11)&"' size='6' /><br>"&BBS.Info(120)&"：<input name='s12' type='text' class='text' value='"&Strings(12)&"' size='6' /><br>"&BBS.Info(122)&"：<input name='s13' type='text' class='text' value='"&Strings(13)&"' size='6' />",90,2
	DIVTR"上传设置：","当论坛系统默认禁止上传后，将不起作用。",GetRadio("s14","禁止",Strings(14),0)&GetRadio("s14","全部会员",Strings(14),1)&GetRadio("s14","只有VIP会员",Strings(14),2),40,1
	Response.Write"<div class='bottom'><input type='submit' value='提 交' class='button' /><input type='reset' value=' 重 置 ' class='button' /></div></div></form>"
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
		If Not Isnumeric(strings(i)) Then GoBack"","一些项目必需用数字填写！":Exit Sub
	Next
	'如果设置限制资源
	For I=10 To 13
		If Int(Strings(I))>0 Then
			Strings(9)=1
			Exit For
		End If
	Next
	
	If Not isnumeric(NewParentID) or BoardName="" Or Introduce="" Then
		GoBack"","":Exit Sub
	ElseIF BoardID=NewParentID Then
		GoBack"","所属论坛不能指定自己！":Exit Sub
	End If
	Set Rs=BBS.execute("Select RootID,Depth,Child,Orders,ParentID,ParentStr From[Board] where BoardID="&BoardID)
	IF Rs.Eof or Rs.Bof Then
		GoBack"系统出错！","该版面不存在，可能已经删除了！"
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
		GoBack"系统出错！","分类不能设置":Exit Sub
	ElseIf Int(NewParentID)<>Int(ParentID) Then
		'判断所指定的论坛是否其下属论坛
		Set Rs=BBS.Execute("select BoardID from [board] where ParentStr like '%"&ParentStr&","&BoardID&"%' and BoardID="&NewParentID)
		if not (Rs.eof and Rs.bof) then
			GoBack"","您不能指定该版面的下属子论坛作为所属论坛"
			Exit sub
		End if
		Rs.Close
		'获得新选的父级
		Set P_rs=BBS.Execute("select * from [board] where Boardid="&NewParentID)
			If P_rs("Depth")+1> MaxDepth Or (Child>0 And P_Rs("Depth")+2>MaxDepth) Then
			GoBack "","本论坛限制了最多只能有" & MaxDepth & "级论坛。如果想使用更多级论坛，请到BBS官方论坛寻求帮助！"
			P_rs.Close
			Set P_rs=Nothing
			Exit Sub
		End If		
	End if
	BBS.Execute("Update [Board] Set BoardName='"&BoardName&"',Strings='"&Join(Strings,"|")&"',Introduce='"&Introduce&"',BoardImg='"&BoardImg&"' where BoardID="&BoardID&"")
  If Int(NewParentID)<>Int(ParentID) Then
	'将一个分论坛移动到其他分论坛下
	'获得所指定的论坛的相关信息
	'得到其下属版面数
	ParentStr=ParentStr & ","
	BoardNum=BBS.Execute("select count(*) from [Board] where ParentStr like '%"&ParentStr & BoardID&"%'")(0)
	If Isnull(BoardNum) Then BoardNum=1
	'在获得移动过来的版面数后更新排序在指定论坛之后的论坛排序数据
	BBS.Execute("update [Board] set orders=Orders + "&BoardNum&"+1  where RootID="&P_rs("RootID")&" And orders>"&P_rs("orders")&"")
	'更新当前版面数据
	If P_rs("parentstr")="0" Then
	BBS.Execute("update [Board] set Depth="&P_Rs("Depth")&"+1,orders="&P_Rs("orders")&"+1,rootid="&P_rs("Rootid")&",ParentID="&NewParentID&",ParentStr='" & P_Rs("BoardID") & "' Where BoardID="&BoardID)
	Else
	BBS.Execute("update [Board] set Depth="&P_Rs("Depth")&"+1,orders="&P_Rs("orders")&"+1,rootid="&P_rs("Rootid")&",ParentID="&NewParentID&",ParentStr='" & P_Rs("ParentStr") & ","& P_Rs("BoardID") &"' Where BoardID="&BoardID)
	End If
	Dim TempParentStr
	i=1
	'更新下属，同时获得移动总数i
	'如果有则更新下属版面数据
	'深度为原有深度加上当前所属论坛的深度
	Set Rs=BBS.Execute("select * from [Board] where ParentStr like '%"&ParentStr & BoardID&"%' order by orders")
	Do while not rs.eof
	i=i+1
	If P_rs("parentstr")="0" Then'如果其父级为类，那么其下属的版面数据
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
	If RootID=P_rs("RootID") then'在同一分类下移动
		'更新所指向的上级论坛版面数，i为本次移动过来的版面数
		'更新其父类版面数
		BBS.Execute("update [Board] set Child=child+"&i&" where (not ParentID=0) and BoardID="&TempParentID)
		For II=1 to P_Rs("depth")
			'得到其父类的父类的版面ID
			Set Rs=BBS.Execute("Select ParentID from [Board] where (not ParentID=0) and BoardID="&TempParentID)
			If Not (rs.eof and rs.bof) then
				TempParentid=Rs(0)
				'更新其父类的父类版面数
			BBS.Execute("update [Board] set Child=child+"&i&" where (not ParentID=0) and BoardID="&TempParentID)
			End if
		Next
		'更新其原父类版面数
		BBS.Execute("update [Board] set Child=child-"&i&" where (not ParentID=0) and BoardID="&ParentID)
		'更新其原来所属论坛数据
		For II=1 to Depth
			'得到其原父类的父类的版面ID
			Set rs=BBS.Execute("select ParentID from [Board] where (not ParentID=0) and BoardID="&ParentID)
			if not (rs.eof and rs.bof) then
				ParentID=rs(0)
				'更新其原父类的父类版面数
				BBS.Execute("update [Board] set child=child-"&i&" where (not ParentID=0) and  boardid="&ParentID)
			End IF
		Next
	Else
	'更新所指向的上级论坛版面数，i为本次移动过来的版面数
	'更新其父类版面数
		BBS.Execute("update [Board] set Child=child+"&i&" where BoardID="&TempParentID)
		For II=1 to P_Rs("depth")
			'得到其父类的父类的版面ID
			Set Rs=BBS.Execute("Select ParentID from [Board] where BoardID="&TempParentID)
			If Not (rs.eof and rs.bof) then
				TempParentid=Rs(0)
				'更新其父类的父类版面数
			BBS.Execute("update [Board] set Child=child+"&i&" where  BoardID="&TempParentID)
			End if
		Next
	'更新其原父类版面数
	BBS.Execute("update [Board] set Child=child-"&i&" where BoardID="&ParentID)
	'更新其原父类的其它版面排序
	BBS.Execute("Update [Board] Set Orders=Orders-"&i&" where RootID="&RootID&" and Orders>"&Orders)
	'更新其原来所属论坛数据
		For II=1 to Depth
			'得到其原父类的父类的版面ID
			Set rs=BBS.Execute("select ParentID from [Board] where BoardID="&ParentID)
			if not (rs.eof and rs.bof) then
				ParentID=rs(0)
				'更新其原父类的父类版面数
				BBS.Execute("update [Board] set child=child-"&i&" where boardid="&ParentID)
			End IF
		Next
	End if
	P_rs.Close:Set P_rs=Nothing
  End If
	Suc"","论坛修改成功 !","Admin_Board.asp"
	BBS.NetLog"操作后台_论坛修改成功!"
	BBS.Cache.clean("BoardInfo")
End Sub


Sub ClassOrders
	Dim BoardID
	Set Rs=BBS.Execute("Select BoardID,BoardName,RootID from[Board] where Depth=0 order by RootID")
	If Rs.Eof Then
		GoBack"","论坛没有分类！请先<a href='?Action=AddClass'> 添加分类</a>"
		Exit Sub
	End If
	Response.Write"<form method=POST style='margin:0' action='?Action=SaveClassOrders'><div class='mian'><div class='top'>分类排序 </div><div class='divth'>排序方式按从小到大排序，请用数字填写，排序数字不能相同。</div>"
	Do while not rs.eof
	DIVTR Rs(1),"","<input name='BoardID' type='hidden' value='"&Rs(0)&"'><input name='RootID' type='hidden' value='"&Rs(2)&"'><input type=text name='NewRootID' value='"&Rs(2)&"' size='4' />",25,1
	Rs.MoveNext
	Loop
	Response.Write"<div class='bottom'><input type='submit' value='修 改' class='button'><input type='reset' value='重 置' class='button'></div></div></form>"
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
			GoBack "排序错误","各分类排序的数字不能一样!"
			Exit Sub
		End If
		Temp=Temp&NewRootID&","
		IF Not IsNumeric(BoardID) or Not isnumeric(NewRootID) Then
			GoBack "排序错误","请用数字填写!"
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
	Suc"","分类排序成功！","?"
	BBS.NetLog"操作后台_分类排序成功！"	
	BBS.Cache.clean("BoardInfo")
End Sub


Sub BoardUnite
	Response.Write"<form method=POST style='margin:0' action='?Action=SaveBoardUnite'><div class='mian'><div class='top'>论坛合并</div>"&_
	"<div class='divth' style='padding:2px;'>将论坛： <select size='1' name='BoardID'><option value=''>请选择原论坛</option>"&BBS.BoardIDList(0,0)&"</select> 合并到论坛： <select size='1' name='NewBoardID'><option value=''>请选择目标论坛</option>"&BBS.BoardIDList(0,0)&"</select> 中 <input type='button' class='button' onclick=""if(confirm('操作后将不能恢复！您确定要合并吗？'))form.submit()"" value='论坛合并'></div>"&_
	"<div class='divtr1' style='padding:5px;'><b>注意事项：</b><font color=red>此操作不可恢复，请慎重操作！</font><br>分类不能操作，不能和其属下的论坛合并。<br>合并后原论坛(包括属下论坛)将被删除，所有帖子(包括属下论坛的帖子)将转移到指定的目标论坛中 </div></div></form>"
End Sub

Sub SaveBoardUnite
	Dim BoardID,NewBoardID,TempParentStr,TempParentID,Rs1,S
	Dim I,AllTable
	Dim ParentStr,Depth,ParentID,Child,RootID
	BoardID=BBS.Fun.Getstr("BoardID")
	NewBoardID=BBS.Fun.Getstr("NewBoardID")
	IF BoardID="" Or NewBoardID="" Then
		GoBack"","请先指定论坛后再进行合并！"
		Exit Sub
	ElseIf BoardID=NewBoardID Then
		Goback"","同一个论坛不用合并了！"
		Exit sub
	End If

	Set Rs=BBS.Execute("Select ParentStr,BoardID,Depth,ParentID,Child,RootID from [board] where BoardID="&BoardID)
	If Rs(2)="0" then
		Goback"系统错误","分类不能做合并操作！"
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
	'判断是否合并到下属论坛
	Set Rs=BBS.Execute("Select BoardID From [Board] where BoardID="&NewBoardID&" And ParentStr like '%"&ParentStr&"%'")
	If Not (rs.eof and rs.bof) then
		Goback"","不能将论坛合并到其下属论坛中!"
		Exit Sub
	End if
	Rs.Close
	'得到全部下属论坛ID
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
	'更新其原来所属论坛版面数
	BBS.Execute("update [Board] set Child=Child-"&child&" where BoardID="&TempParentID)
	'更新其原来所属论坛数据，排序相当于剪枝而不需考虑
	For I=1 to Depth
		'得到其父类的父类的版面ID
		Set rs=BBS.Execute("select ParentID from [Board] where boardID="&TempParentID)
		If Not (rs.eof and rs.bof) then
			TempParentID=rs(0)
			BBS.Execute("update [Board] set Child=Child-"&Child&" where boardid="&TempParentID)
		End if
	Next
	'更新论坛帖子数据
	AllTable=Split(BBS.BBStable(0),",")
	For i=0 To uBound(AllTable)
		BBS.Execute("update [BBS"&AllTable(i)&"] set BoardID="&NewBoardID&" where BoardID in ("&ParentStr&")")
	Next
	BBS.Execute("update [Topic] set BoardID="&NewBoardID&" where BoardID in ("&ParentStr&")")
	'删除被合并论坛
	Set Rs=BBS.Execute("Select Sum(EssayNum),sum(TopicNum),sum(TodayNum) from [Board] where RootID="&RootID&" And BoardID in ("&ParentStr&")")
	BBS.Execute("Delete from [Board] where RootID="&RootID&" And BoardID in ("&ParentStr&")")
	'更新新论坛帖子计数
	BBS.Execute("update [Board] set EssayNum=EssayNum+"&rs(0)&",TopicNum=TopicNum+"&rs(1)&",TodayNum=TodayNum+"&rs(2)&" where BoardID ="&NewBoardID&"")
	'更新上级版块
	set Rs1=BBS.Execute("select Depth,ParentStr,Boardid from [Board] where BoardID="&NewBoardID)
	If Rs1(0)>1 Then
	ParentStr=Rs1(0)
	BBS.Execute("update [Board] set EssayNum=EssayNum+"&rs(0)&",TopicNum=TopicNum+"&rs(1)&",TodayNum=TodayNum+"&rs(2)&" where boardid in ("&ParentStr&")")
	End If
	Rs1.Close:Set Rs1=Nothing
	Rs.Close
	'调整版主
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
	Suc"","合并成功！已经将原论坛（包括属下）的所有帖子合并到目标论坛。","?"
	BBS.NetLog"操作后台_合并论坛成功!"
	'更新版块缓存
	BBS.Cache.clean("BoardInfo")
	BBS.Cache.clean("Board"&NewBoardID)
	BBS.Cache.clean("Board"&BoardID)
End Sub


Sub ClearData
	Set Rs=BBS.execute("Select BoardName,TopicNum,EssayNum From[Board] where BoardID="&BBS.BoardID&"")
	IF Rs.Eof Then
		GoBack"","论坛版面不存在，可能经被删除了"
		Exit Sub
	End If
	Response.Write"<div class='mian'><div class='top'>"&Rs("BoardName")&" 数据清理</div>"&_
	"<div class='divtr1' style='padding:5px;'><b>本版的帖子信息</b><br />主题数：<span style='color:#F00'>"&Rs("TopicNum")&"</span>&nbsp;&nbsp;总帖数：<span style='color:#F00'>"&Rs("EssayNum")&"</span>&nbsp;&nbsp;精华主题数：<span style='color:#F00'>"&BBS.Execute("Select Count(TopicID) From[Topic] where IsGood=1 and BoardID="&BBS.BoardID&"")(0)&"</span></div>"&_
	"<div class='divth' style='padding:3px'><form method=POST style='margin:0' action='?Action=StartClearData'>清除 <b>"&Rs(0)&"</b> 在 <input name='BoardID' value='"&BBS.BoardID&"' type='hidden' /><select name='SqlTableID'><option value='0'>所有数据表</option>"&SqlTableList&"</select> 中 <input type='text' class='text' name='ClearDate' value='365' size='5'> 天前的帖子。 <input type='submit' value='执行清理' class='button'></form></div>"&_
	"<div class='divtr2' style='padding:5px;'><b>注意事项</b><br><font color=red>此操作不可恢复！精华帖子不会被删除！</Font><br>如果您的论坛数据众多，执行此操作将消耗大量的服务器资源。<br>执行过程请耐心等候，最好选择夜间在线人少的时候更新。</div></div>"
	Rs.Close
End Sub

Sub StartClearData
	Dim SqlTableID,ClearDate,BoardID,AllTable,i,Temp
	SqlTableID=request.form("SqlTableID")
	ClearDate=request.form("ClearDate")
	BoardID=request.form("BoardID")
	If Not isnumeric(ClearDate) or Not isNumeric(SqlTableID) or Not isNumeric(BoardID)  Then
		GoBack"","请用数字填写！"
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
		IF Rs(1)=1 Then'删除投票
			BBS.Execute("Delete from [TopicVote] where TopicID="&RS(0)&"")
			BBS.Execute("Delete from [TopicVoteUser] where TopicID="&RS(0)&"")
		End If
		Rs.movenext
		Loop
		Rs.Close
		BBS.Execute("Delete From[Topic] where BoardID="&BoardID&" And SqlTableID="&AllTable(i)&" And IsGood=0 And  DATEDIFF('d',[LastTime],'"&BBS.NowBbsTime&"')>"&ClearDate&" ")
	Next
	Temp=BBS.Execute("Select BoardName From[Board] where BoardID="&BoardID)(0)
	BBS.NetLog"操作后台_清理论坛版面："&Temp&" 在"&ClearDate&"天前的数据成功!"
	Suc"","成功的清理了论坛数据！建议做一下<a href='?Action=BoardUpdate'>论坛整理</a>","?"
End Sub


Sub PassUser
	Dim Temp
	Set Rs=BBS.execute("Select PassUser,BoardName,Strings From [Board] where BoardID="&BBS.BoardID&" And ParentID<>0")
	IF Rs.eof Then
		GoBack"","此论坛的类型不是认证论坛，不能设置认证用户。"
		Exit Sub
	End If
	Temp=Split(Rs(2),"|")
	If Temp(6)="0" Then GoBack"","此论坛的类型不是认证论坛，不能设置认证用户。":Exit Sub
	Response.Write"<form method=POST style='margin:0' action='?Action=SavePassUser'><div class='mian'><div class='top'>修改论坛论证用户</div>"
	DIVTR"所在论坛：","","<input name='BoardID' value='"&BBS.BoardID&"' type='hidden'>"&Rs("BoardName"),25,1
	DIVTR"通过认证的用户：","各用户之间用“|”隔开<br />请不要使用回车键Enter","<textarea name='PassUser' rows='5'>"&Rs("PassUser")&"</textarea>",70,2
	Response.Write"<div class='bottom'><input type='submit' value='提 交' class='button' /><input type='reset' class='button' value='重 置' /></div></div></form>"
	Rs.Close
End Sub

Sub SavePassUser
	Dim PassUsers,BoardID
	BoardID=BBS.Fun.GetStr("BoardID")
	PassUsers=Trim(Replace(Request.Form("PassUser"),"'",""))
	PassUsers=Replace(PassUsers,chr(10), "")
	PassUsers=Replace(PassUsers,chr(13), "")
	BBS.Execute("Update [Board] Set PassUser='"&PassUsers&"' where BoardID="&BoardID&" And ParentID<>0")
	BBS.NetLog"操作后台_更新认证会员成功!"
	Suc"","成功的更新了该论坛的认证会员！","?"
	BBS.Cache.clean("BoardInfo")
End Sub


Sub BoardUpdate
	Response.Write"<div class='mian'>"&_
	"<div class='divth' style='height:50px'><b><div id='BBST'>数据版面正在整理，请稍等</div></b><div style='margin:2px auto 0;width:400px;height:16px;background:#DEFAF1;text-align:left'><img src='Images/icon/hr1.gif' width=0 height='16' id='BBSimg' align='absmiddle' alt='进度条' /></div>"&_
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
	LastCaption="无"
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
	'如果有上级论坛，那么更新上级论坛
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
		BoardAdmin="无"
	Else
		BoardAdmin=Replace(Boardadmin,"|","、")
	End If
'更新缓存
	BBS.Cache.clean("Board"&SQL(0,i))
	Table "论坛 <Font color=blue>"&SQL(1,i)&"</Font> 整理成功","总帖数"&EssayNum&" | 主题数："&TopicNum&" | 今日帖："&TodayNum&" | 版主："&BoardAdmin&" | 最新主题："&LastCaption&""
	II=II+1
	Response.Write "<script>document.getElementById(""BBSimg"").style.width=" & Fix((ii/BoardNum) * 400) & ";" & VbCrLf
	Response.Write "document.getElementById(""BBStxt"").innerHTML=""" & FormatNumber(ii/BoardNum*100,4,-1) & """;" & VbCrLf
	Response.Write "</script>" & VbCrLf
	Response.Flush
	Next
	End If
	Response.Write "<script>document.getElementById(""BBSimg"").style.width=400;document.getElementById(""BBStxt"").innerHTML=""100"";document.getElementById(""BBST"").innerHTML=""<font color=red>成功完成整理！</font>"";</script>"
	BBS.NetLog"操作后台_整理论坛!"
End Sub

Function SqlTableList()
	Dim AllTable,I
	AllTable=Split(BBS.BBStable(0),",")
	For i=0 To uBound(AllTable)
	SqlTableList=SqlTableList&"<option value='"&AllTable(I)&"'>数据表"&AllTable(I)&"</option>"
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
		GoBack"系统出错！","该版面不存在，可能已经删除了！"
		Exit Sub
	End If
	Orders=Rs(0)
	ParentID=Rs(1)
	ParentStr=Rs(2)
	Rs.Close
	'当版面为类时
	If ParentID=0 then GoBack"系统出错！","版面ID出错。":Exit Sub
	'得到其下属版面数
	ParentStr=ParentStr & ","
	BoardNum=BBS.Execute("select count(*) from [Board] where ParentStr like '%"&ParentStr & BoardID&"%'")(0)
	If Isnull(BoardNum) Then BoardNum=1
	'获得父级信息
	Set P_rs=BBS.Execute("select * from [board] where Boardid="&ParentID)
	'在获得移动过来的版面数后更新排序在指定论坛之后的论坛排序数据
	BBS.Execute("update [Board] set orders=Orders + "&BoardNum&"+1  where RootID="&P_rs("RootID")&" And orders>"&P_rs("orders")&"")
	'更新当前版面数据
	BBS.Execute("update [Board] set orders="&P_Rs("orders")&"+1 Where BoardID="&BoardID)
	Dim TempParentStr
	i=1
	'更新下属，同时获得移动总数i
	'如果有则更新下属版面数据
	Set Rs=BBS.Execute("select * from [Board] where ParentStr like '%"&ParentStr & BoardID&"%' order by orders")
	Do while not rs.eof
	i=i+1
	If P_rs("parentstr")="0" Then'如果其父级为类，那么其下属的版面数据
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
	BBS.NetLog"操作后台_调整论坛排序成功!"
	Response.Redirect"?"
	Response.End
End Sub
%>