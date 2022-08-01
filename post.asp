<!--#include file="inc.asp"--><%
Dim Action,ID,Page,Temp
Dim Caption,SubmitUrl
Dim Title,Content
If Not BBS.Founduser Then BBS.GoToerr(31)
BBS.CheckBoard()
ID=BBS.CheckNum(request.querystring("ID"))
Page=BBS.CheckNum(request.querystring("page"))
Action=lcase(request.querystring("action"))
If Len(Action)>10 Then BBS.GoToerr(1)
If Session(CacheName & "MyGradeInfo")(10)="1" Then 
	Temp=" <select name='font_s'><option value=''>样式</option><option value='B'>粗体</option><option value='I'>斜体</option><option value='U'>加线</option><option value='no'>默认</option></select><select name='font_c'><option value=''>颜色</option><option style='COLOR:#000;BACKGROUND-COLOR:#000' value='#000'></option><option style='COLOR:#000;BACKGROUND-COLOR:#F00' value='#F00'></option><option style='COLOR:#000;BACKGROUND-COLOR:#00F' value='#00F'></option><option style='COLOR:#000;BACKGROUND-COLOR:#0F0' value='#0F0'></option><option style='COLOR:#000;BACKGROUND-COLOR:#008000' value='#008000'></option><option style='COLOR:#000;BACKGROUND-COLOR:#FA0' value='#FA0'></option><option style='COLOR:#000;BACKGROUND-COLOR:#F0F' value='#F0F'></option><option style='COLOR:#000;BACKGROUND-COLOR:#0FF' value='#0FF'></option><option style='COLOR:#000;BACKGROUND-COLOR:#888' value='#888'></option><option style='COLOR:#000;BACKGROUND-COLOR:#800000' value='#800000'></option><option style='COLOR:#000;BACKGROUND-COLOR:#800080' value='#800080'></option><option style='COLOR:#000;BACKGROUND-COLOR:#008080' value='#008080'></option><option style='COLOR:#000;BACKGROUND-COLOR:#000080' value='#000080'></option><option style='COLOR:#000;BACKGROUND-COLOR:#808000' value='#808000'></option><option value='no'>默认</option></select>"
End If
Title=BBS.Row("<b>帖子主题：</b><select name='Title' onChange='DoTitle(this.options[this.selectedIndex].value)' style='font-size: 9pt'><option selected value=''>话题</option><option value='[原创]'>[原创]</option><option value='[转帖]'>[转帖]</option><option value='[灌水]'>[灌水]</option><option value='[讨论]'>[讨论]</option><option value='[求助]'>[求助]</option><option value='[推荐]'>[推荐]</option><option value='[公告]'>[公告]</option><option value='[注意]'>[注意]</option><option value='[贴图]'>[贴图]</option><option value='[建议]'>[建议]</option><option value='[下载]'>[下载]</option><option value='[分享]'>[分享]</option></select>","<input id='caption' name='caption' type='text' style='width:60%' maxlength='200' />"&Temp,"75%","")

Select Case Action
	Case"vote"
		Vote()
	Case"reply"
		Reply()
	Case"edit"
		Edit()
	Case Else
		BBS.Stats="发表新帖"
		Submiturl="postsave.asp?boardid="&BBS.boardid
End Select
BBS.Head "post.asp?boardid="&BBS.boardid,BBS.BoardName,BBS.Stats
ShowMain()
BBS.Footer()
Set BBS =Nothing

Sub Vote()
	Dim i
	If Session(CacheName & "MyGradeInfo")(12)="0" Then
		Temp="<div style=""padding:4px"">对不起，您目前的论坛等级没有发表投票主题的权限。</div>"
	Else
		Temp="请选择投票项目数：<select name='votenum' id='votenum' onchange='SetNum(this)' />"
		For i = 2 to int(BBS.Info(63))
			Temp=Temp&"<option value='"&I&"'>"&I&"</option>"
		Next
		Temp=Temp&"</select>允许多选<input type='checkbox' name='votetype' value='2' /> 过期时间：<select name='outtime'><option value='1'>一天</option><option value='3'>三天</option><option value='7'>一周</option><option value='15'>半个月</option><option value='31'>一个月</option><option value='93'>三个月</option><option value='365'>一年</option><option value='10000' selected>不过期</option></select><hr size=1 width='98%' /><div id='optionid'><div>选项1：<input type='text' name='Votes1' style='width:80%' /></div><div>选项2：<input type='text' name='votes2' style='width:80%' /><INPUT TYPE='hidden' name='autovalue' value='2' /></div></div>"
	End If	
	Title=Title&BBS.Row("<b>投票选项：</b>",Temp,"75%","")
	BBS.Stats="发表新投票"
	SubmitUrl="postsave.asp?boardid="&BBS.boardid
End Sub

Sub Reply()
	Dim Rs,BbsID
	if ID=0 Then BBS.GoToErr(1)
	BBS.Stats="回复帖子"
	Set Rs=BBS.Execute("Select Caption,SqlTableID,IsLock,IsDel From [Topic] where TopicID="&ID&" And IsDel=0")
	If Rs.Eof Then
		BBS.GoToErr(21)
	ElseIf Rs(2)=1 Then
		BBS.GoToErr(22)
	Else
		Title=BBS.Row("<input type=hidden name='caption' id='caption' value='Re:"&Rs(0)&"' />回复主题：",Rs(0),"75%","22px")
		BBS.TB=Rs(1)
	End If
	Rs.close
	Set Rs=Nothing
	Submiturl="postsave.asp?Action=Reply&boardid="&BBS.boardid&"&TB="&BBS.TB&"&ID="&ID&"&page="&page

	BbsID=BBS.CheckNum(Request.querystring("BbsID"))
	If BbsID>0 Then
		Set Rs=BBS.Execute("select top 1 B.ReplyTopicID,B.TopicID,B.Name,B.AddTime,B.Content,B.boardid,U.IsShow from [Bbs"&BBS.TB&"] As B inner join [User] As U on B.Name=U.Name where B.BbsID="&BbsID&" And B.IsDel=0")
			If Not Rs.Eof Then
				If Rs(1)<>ID And Rs(0)<> ID Then BBS.GoToErr(1)
				If Rs(6)=1 Then
					Content="<div class=""quote"">引用 "&RS(2)&" 的发言内容:<br><font color=""#F00"">屏蔽内容不能引用</font><div><br>"
				Else
				If BBS.Info(60)="0" Then
				Content="<div class=""quote"">以下是引用 <b>"&RS(2)&"</b></font> 在(<i>"&Rs(3)&"</i>)的发言<br>"&QuoteCode(Rs(4))&"</div><br><br>"
				Else
				Content="[quote]以下是引用 [B]"&RS(2)&"[/B] ：<br>"&QuoteCode(Rs(4))&"<br>[/quote]<br>"			
				End If
				End If
			End if
			Rs.close
		Set Rs=Nothing
	End If
End Sub

Sub Edit()
	Dim Rs,BbsID,TopicIsLock,TopicRs,IsTop
	BbsID=BBS.CheckNum(request.querystring("BbsID"))
	IF BbsID=0 Or ID=0 Then BBS.GoToErr(1)
	
	Set Rs=BBS.Execute("Select boardid,TopType,SqlTableID,IsLock From [Topic] where IsDel<>1 And TopicID="&ID)
	If Rs.Eof Then
		BBS.GoToErr(58)
	Else
		TopicRs=Rs.GetRows(-1)
	End If
	Rs.Close
	Set Rs=BBS.Execute("select boardid,Name,AddTime,TopicID,Caption,Content,IsDel From [Bbs"&TopicRs(2,0)&"] where IsDel<>1 And BbsID="&BbsID&"")
	If Rs.eof  Then
		BBS.GoToErr(58)
	Else
	If lcase(BBS.MyName)=lcase(rs("name")) Then 
			If TopicRs(3,0)=1 And BBS.MyAdmin<>9 Then BBS.GoToErr(22)
			If Session(CacheName & "MyGradeInfo")(22)="0" Then
				If BBS.Info(12)<>"0" And DateDiff("s",Rs("AddTime")+BBS.Info(12)/1440,BBS.NowBbsTime)>0 Then BBS.GoToErr(34)
			End If
	Else
		If Session(CacheName & "MyGradeInfo")(24)="0" Then BBS.GoToErr(33)
		If TopicRs(1,0)=5 or TopicRs(1,0)=4 Then'如果是总顶或区顶
			If TopicRs(0,0)<>BBS.boardid Then'如果不是本版，版主无权
				If BBS.MyAdmin=7 Then BBS.GoToErr(51)
			End If
		Else
			If BBS.MyAdmin=7 And Not BBS.IsBoardAdmin Then BBS.GoToErr(71)
		End If
	End If
		If TopicRs(1,0)=5 or TopicRs(1,0)=4 Then
			If lcase(BBS.MyName)<>lcase(rs("name")) Then

			End If
		Else
			If TopicRs(0,0)<>BBS.boardid Then BBS.GotoErr(1)
		End If

		IF Rs("TopicID")=0 Then
			Title=BBS.Row("<input type=hidden id='caption' name='caption' value='"&Rs(4)&"' /><b>编辑回复帖：</b>",rs(4),"75%","23px")
		Else 
			Title=replace(Title,"id='caption'","id='caption' value='"&Rs(4)&"'")
		End IF
		Content=ReplaceUBB(rs(5))
	End if
	Rs.Close
	BBS.Stats="编辑帖子"
	Submiturl="postsave.asp?Action=Edit&ID="&ID&"&BbsID="&BbsID&"&boardid="&BBS.boardid&"&TB="&TopicRs(2,0)&"&page="&page&""
End Sub


Function ShowMain()
	With BBS
	Dim Face,I,Temp1,S1
	Temp="<form style='margin:0;' name='preview' action='preview.asp?Action=preview' method='post' target='preview'><input type='hidden' name='pcaption' /><input type='hidden' name='pcontent' /></form>"
	Temp=Temp&"<form style='margin:0;' method=POST name='say' action='"&Submiturl&"' >"
	Temp=Temp&title
	If .Info(15)="1" Then
		Temp=Temp&.Row("<b>发帖验证码：</b>",.GetiCode,"75%","")
	Else
		Temp=Temp&"<input type=hidden name='iCode' id='iCode' value='BBS' />"
	End If
	Face="<input name=face type=radio value=1 checked class=checkbox /><img src='pic/face/1.gif' border='0' align='absmiddle' atl='' />&nbsp;"
	For i=2 to 18
		Face=Face&"<input type=radio value="&i&" name='face' class=checkbox /><img border=0 align='absmiddle' src='pic/face/"&i&".gif' atl='' />&nbsp;"
		if i=9 then Face=Face&"<br />"
	Next
	Temp=Temp&.Row("<textarea id='content' name='content' style='display:none'>"&Server.HtmlEnCode(Content)&"</textarea><b>你的表情：</b><br />在帖子前面",Face,"75%","")
	If .Info(30)="0" Then
	  Temp1="本论坛暂时关闭上传功能。<br>"
	 ElseIf Session(CacheName & "MyGradeInfo")(14)="0" then
	  Temp1="您目前的论坛等级组没有上传的权限！"
	 ElseIf .BoardString(14)="0" then
	  Temp1="本版面暂时关闭上传功能。"
	 ElseIf .BoardString(14)="2" And Session(CacheName & "MyInfo")(17)="0" then
	  Temp1="本版面只允许VIP会员有上传权限！"
	Else
		Temp1="<input style=""margin-top:10px"" class=""button"" type=""button"" value=""上传附件"" onclick=""javascript:document.getElementById('up').style.display='block';upf.location.replace('UploadFile.asp');this.style.display='none'""> 可上传文件类型："&Replace(.Info(34)&"|"&.Info(35),"|","、")
		Temp1=Temp1&"<div id='up' style='display:none'><iframe id='upf' name='upf' scrolling='no' frameborder='0' height='22' width='100%'></iframe></div>"
	End if
	Temp=Temp&.Row("<b>附件上传：</b><br />每日您可以上传<font color=blue>"&Session(CacheName & "MyGradeInfo")(15)&"</font>个(最大<font color=blue>"&Session(CacheName & "MyGradeInfo")(16)&"</font>KB)",Temp1,"75%","42px")	
	Temp1="<br /><a href=""javascript:CheckLength("&Session(CacheName & "MyGradeInfo")(9)&")""> 内容限制：<font color=red>"&Session(CacheName & "MyGradeInfo")(9)&"字节</font></a><br />HTML标签：<font color=red>"
	If .Info(60)="1" Then Temp1=Temp1&"×" Else Temp1=Temp1&"√"
	Temp1=Temp1&"</font><br />UBB标签： <font color=red>√</font><br />上传文件：<font color=red>"
	If .Info(30)="0" Then Temp1=Temp1&"×" Else Temp1=Temp1&"√"
	Temp1=Temp1&"</font><br /><b>发特殊帖：</b><br />"&_
	Especial("回复可见","Especial('[REPLY]','[\/REPLY]')",.Info(70))&_
	Especial("金钱可见","Coin()",.Info(71))&"<br />"&_
	Especial("积分可见","Mark()",.Info(72))&_
	Especial("日期可见","Showdate()",.Info(73))&"<br />"&_
	Especial("性别可见","Sex()",.Info(74))&_
	Especial("登陆可见","Especial('[LOGIN]','[\/LOGIN]')",.Info(75))&"<br />"&_
	Especial("指定读者","Name()",.Info(76))&_
	Especial("付费观看","Buypost()",.Info(77))&"<br />"&_
	Especial("插入代码识别转换","Code()",.Info(68))
	Temp1=Temp1&"<ul><li>发帖请遵守国家法律</li><li>禁止发表政治及色情内容</li></ul>"
	If .Info(60)="1" Then S1="UbbEdit()" Else S1="HtmlEdit()"
	S1="<script type=""text/javascript"">"&S1&"</script>"
	If BBS.CC(0)="1" Then
		S1=S1&"<object width='72' height='24'><param name='wmode' value='transparent' /><param name='allowScriptAccess' value='always' /><param name='movie' value='http://union.bokecc.com/flash/plugin_"&BBS.CC(4)&".swf?userID="&BBS.CC(2)&"&type=BBS' /><embed src='http://union.bokecc.com/flash/plugin_"&BBS.CC(4)&".swf?userID="&BBS.CC(2)&"&type=BBS' type='application/x-shockwave-flash' width='72' height='24' allowScriptAccess='always'></embed></object> &nbsp;"
	End IF
	If Action="edit" And Session(CacheName & "MyGradeInfo")(25)="1" Then S1= S1 & "<input name='editchalk' type='checkbox' value='No' class=checkbox>不留下编辑标记"
	If Action="" then
	If (BBS.MyAdmin=7 And Not BBS.IsBoardAdmin) Then
	Else
	S1=S1&"主题设置："
	If Session(CacheName&"MyGradeInfo")(31)="1" Then S1= S1 & "<input name='top' type='checkbox' value='1' class=checkbox>置顶 "
	If Session(CacheName&"MyGradeInfo")(32)="1" Then S1= S1 & "<input name='classtop' type='checkbox' value='1' class=checkbox>区置顶 "
	If Session(CacheName&"MyGradeInfo")(33)="1" Then S1= S1 & "<input name='alltop' type='checkbox' value='1' class=checkbox>总置顶 "
	If Session(CacheName&"MyGradeInfo")(34)="1" Then S1= S1 & "<input name='good' type='checkbox' value='1' class=checkbox>精华 "
	If Session(CacheName&"MyGradeInfo")(35)="1" Then S1= S1 & "<input name='lock' type='checkbox' value='1' class=checkbox>锁定"
	End If
	End If
	Temp=Temp&.Row("<b>帖子内容：</b>"&Temp1,S1,"75%","")
	Temp=Temp&"<div align='center' style="" padding:5px;BACKGROUND: "&.SkinsPIC(1)&";"">"
	Temp=Temp&"&nbsp;<input type='button' value='OK 发表' id='sayb' onclick='checkform("&Session(CacheName & "MyGradeInfo")(9)&")' class='button' /> <input type=button value='预 览' onclick='Gopreview()' class='button' /> <input type='reset' value='NO 重写' onclick='Goreset()' class='button' />" 
	Temp=Temp&"</div></form>"
	.ShowTable .Stats,Temp
	End With
End Function

Function replaceUBB(str)
	dim re
	If Str="" Then Exit Function
	Set re=new RegExp
	re.IgnoreCase=true
	re.Global=True
	re.Pattern="(>)("&vbNewLine&")(<)"
	Str=re.Replace(Str,"$1$3")
	re.Pattern="(>)("&vbNewLine&vbNewLine&")(<)"
	Str=re.Replace(Str,"$1$3")
	re.Pattern=vbNewLine
	Str=re.Replace(Str,"<br>")	
	re.Pattern="(\[right\])(\[color=(.[^\[]*)\])(.[^\[]*)(\[\/color\])(\[\/right\])"
	str=re.Replace(str," ")
	re.Pattern="(<div style=""color:#999999;text-align:right"">「该帖子被(.*)编辑过」<\/div>)"
	str=re.Replace(str," ")
	str=Replace(Str,"  ","&nbsp;&nbsp;")
	Set re=Nothing
	replaceUBB=str
End function

Function Especial(eName,gourl,Flag)
	If flag="1" Then
		Especial="<a href=""javascript:"&Gourl&""">"&eName&"</a> <span style='color:#F00'>√</span> "
	Else
		Especial=eName&" <span style='color:#AAA'>×</span> "
	End If
End Function

Function QuoteCode(str)
		Dim re,restr
		Set re=new RegExp
		re.IgnoreCase=true
		re.Global=True
		restr="<hr>加密内容不能引用 <hr>"
		re.Pattern="(\[DATE=(.[^\[]*)\])(.+?)(\[\/DATE\])"
		str=re.Replace(str,restr)
		re.Pattern="(\[SEX=*([0-1]*)\])(.+?)(\[\/SEX\])"
		str=re.Replace(str,restr)
		re.Pattern="(\[COIN=*([0-9]*)\])(.+?)(\[\/COIN\])"
		str=re.Replace(str,restr)		
		re.Pattern="(\[USERNAME=(.[^\[]*)\])(.+?)(\[\/USERNAME\])"
		str=re.Replace(str,restr)	
		re.Pattern="(\[GRADE=*([0-9]*)\])(.+?)(\[\/GRADE\])"
		str=re.Replace(str,restr)	
		re.Pattern="(\[MARK=*([0-9]*)\])(.+?)(\[\/MARK\])"		
		str=re.Replace(str,restr)
		re.Pattern="(\[BUYPOST=*([0-9]*)\])(.+?)(\[\/BUYPOST\])"
		str=re.Replace(str,restr)
		re.Pattern=vbcrlf&vbcrlf&vbcrlf&"(\[RIGHT\])(\[COLOR=(.[^\[]*)\])(.[^\[]*)(\[\/COLOR\])(\[\/RIGHT\])"
		str=re.Replace(str,"")
		re.Pattern="(\[reply\])(.+?)(\[\/reply\])"
		Str=re.Replace(str,restr)	
		QuoteCode=replaceUBB(str)
		Set re=Nothing
End Function
%>