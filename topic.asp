<!--#include file="Inc.asp"-->
<!--#include file="inc/ubb_Cls.asp"-->
<!--#include file="inc/page_Cls.asp"-->
<%
Dim TopicCaption,TopicUserName,TopicTopType,TopicIsGood,TopicIsLock,TopicIsVote,TopicHits,TopicReplyNum
Dim Page,PageString,PageInfo,ID,UserName,BBSID
ID=BBS.CheckNum(request.querystring("ID"))
IF BBS.boardid=0 or ID=0 or BBS.TB=0 then BBS.GoToErr(1)
BBS.CheckBoard()
GetTopicInfo()
Show()
FastReply()
BBS.Footer()
BBS.Execute("UPDATE [Topic] SET Hits=Hits+1 WHERE TopicID="&id&"")
Set BBS =Nothing
	
Sub GetTopicInfo()
	Dim Rs,Arr_Rs,PageUrl
	Set Rs=BBS.Execute("Select TopicID,Caption,Name,TopType,IsGood,IsLock,isVote,Hits,ReplyNum,Face,AddTime From [Topic] where TopicID="&id&" And (boardid="&BBS.boardid&" or TopType=4 or TopType=5) and IsDel=0")
	IF Rs.eof then BBS.GoToErr(21)
	Arr_Rs=Rs.GetRows(1)
	Rs.Close
	Set Rs=Nothing
	TopicCaption =BBS.Fun.HtmlCode(Arr_Rs(1,0))
	TopicUserName=Arr_Rs(2,0)
	TopicTopType =Arr_Rs(3,0)
	TopicIsGood  =Arr_Rs(4,0)
	TopicIsLock  =Arr_Rs(5,0)
	TopicIsVote  =Arr_Rs(6,0)
	TopicHits    =Arr_Rs(7,0)
	TopicReplyNum=Arr_Rs(8,0)
	IF TopicIsGood=1 Then If BBS.Info(64)="0" And Not BBS.FoundUser Then BBS.GoToErr(25)
	If Request.QueryString("page") > 1 Then
	  PageUrl = "&Page="&Request.QueryString("page")
	Else
	  PageUrl = ""
	End If
	IF BBS.BoardString(5)="1" or BBS.BoardString(6)="1" Then
		BBS.Head "topic.asp?id="&id&"&boardid="&BBS.boardid&"&tb="&BBS.TB&PageUrl,BBS.Boardname,"浏览保密帖子"'屏蔽信息
	Else
		BBS.Head "topic.asp?id="&id&"&boardid="&BBS.boardid&"&tb="&BBS.TB&PageUrl,BBS.Boardname,BBS.Fun.StrLeft(Arr_Rs(1,0),40)
	End IF
End Sub

Function SayBar()
	If BBS.BoardString(0)="0" or BBS.MyAdmin=9 or BBS.MyAdmin=8 or (BBS.MyAdmin=7 And BBS.IsBoardAdmin) Then
		SayBar="<a href='post.asp?boardid="&BBS.boardid&"'>"&BBS.SkinsPIC(7)&"</a> <a href='post.asp?action=vote&boardid="&BBS.boardid&"'>"&BBS.SkinsPIC(8)&"</a>"
		If TopicIsLock=0 or BBS.MyAdmin=9 Then
			SayBar=SayBar&" <a href='post.asp?action=reply&boardid="&BBS.boardid&"&ID="&id&"'>"&BBS.SkinsPIC(9)&"</a>"
		End If
	End If
End Function

Function Show()
	Dim S,S1
	S=BBS.ReadSkins("帖子表格")
	S=Replace(S,"{显示投票}",ShowVote())
	S=Replace(S,"{发帖按钮}",SayBar())
	S=Replace(S,"{点击数}",TopicHits+1)
	S=Replace(S,"{主题管理}",SetTopic())
	S=Replace(S,"{标题}",TopicCaption)
	S1=ShowBBS()
	S=Replace(S,"{分页}",PageInfo)
	S=Replace(S,"{显示帖子}",S1)
	S=Replace(S,"{版块下拉列表}","<script language=""JavaScript"" type=""text/javascript"">BoardSelect()</script>")
	Response.Write(S)
End Function

Function TopicMood()
	Dim S
	IF TopicIsGood=1 Then S=BBS.SkinsPIC(13)&" 精华帖子 "
	IF TopicIsVote=1 then S=S&BBS.SkinsPIC(14)&" 投票帖子 "
	IF TopicIsLock=1 then S=S&BBS.SkinsPIC(17)&" 此帖锁结 "
	IF TopicTopType=3 then S=S&BBS.SkinsPIC(12)&" 置顶帖子 "
	IF TopicTopType=4 then S=S&BBS.SkinsPIC(11)&" 区置顶帖子 "
	IF TopicTopType=5 then S=S&BBS.SkinsPIC(10)&" 总置顶帖子"
	IF S<>"" Then S="<div class=""topicmood"" style=""float:right;"">"&S&"</div>"
	TopicMood=S
End Function

Function ShowVote()
	If TopicIsVote=0 Then Exit Function
	Dim S,Rs,Arr_Rs,Vote,VoteNum,AllvoteNum,VotePicW,Opt,ClueTxt,CanVote,VoteShow,i,ii
	Set Rs=BBS.Execute("Select TopicID,Vote,VoteNum,VoteType,OutTime From [TopicVote] where TopicID="&id&"")
	If Rs.Eof then Exit Function
	Arr_Rs=Rs.GetRows(1)
	Rs.Close:Set Rs=Nothing
	Vote=Split(Arr_Rs(1,0),"|")
	VoteNum=split(Arr_Rs(2,0),"|")
	CanVote=True
	If Not BBS.FoundUser Then
		ClueTxt="您还没有登陆，不能进行投票。"
		CanVote=False
	Else
		S="【<a style='cursor:pointer;color:#F00;' onclick=""openwin('preview.asp?action=vote&Id="&Arr_Rs(0,0)&"',550,400,'yes')"">详情</a>】"
		If Session(CacheName & "MyGradeInfo")(11)="0" Then
			ClueTxt="您的等级没有参加投票的权限。"
			CanVote=False
		End If
		
		IF not BBS.Execute("Select User From [TopicVoteUser] where User='"&BBS.MyName&"' and TopicID="&id&"").eof then
			ClueTxt="您已经投过票了，不能再投票了。"&S
			CanVote=False
		ElseIf SESSION(CacheName& "MyGradeInfo")(37)="1" Then
			ClueTxt=ClueTxt&S
		End If
		If SESSION(CacheName& "MyGradeInfo")(38)="1" Then ClueTxt=ClueTxt&" 【<a href=settopic.asp?action=editvote&TB="&BBS.TB&"&boardid="&BBS.boardid&"&ID="&id&">修改投票选项</a>】"
	End If
	
	For i=1 to ubound(Vote)
		AllvoteNum=Int(AllvoteNum+VoteNum(i))
	Next
	IF AllVoteNum=0 then AllvoteNum=1
	For i=1 To ubound(Vote)
		ii=ii+1
		VotePicW=VoteNum(i)/AllvoteNum*85
		If CanVote Then
			IF Int(Arr_Rs(3,0))=1 then Opt="<input class=checkbox type='radio' value='"&i&"' name='opt' />" Else Opt=" <input class=checkbox type='checkbox' name='opt"&i&"' value='"&i&"' />"
		End If
		IF ii>6 Then ii=1
		VoteShow=VoteShow&"<div style='height:25px'><div style='text-align:left;width:50%;float:right;' ><img height='8' width='"&VotePicW&"%' src='Images/hr"&ii&".gif'> <strong>"&VoteNum(i)&"</strong> 票</div><div style='text-align:left;'>"&i&"."&Opt&BBS.Fun.HtmlCode(Vote(i))&"</div></div>"
	Next
	VoteShow="<div >"&VoteShow&"</div>"
	
	If DateDiff("s",BBS.NowBbsTime,Arr_Rs(4,0))<0 then
		ClueTxt=ClueTxt&"该投票已经过期，不能进行投票。"
		CanVote=False
	End If
	IF CanVote then
		ClueTxt="<input type='submit' class=button value='投 票 (投票后方能看详尽结果)'>"&ClueTxt
	End IF
	ClueTxt=ClueTxt&" [ 截止时间："&Arr_Rs(4,0)&" ]"

	S=BBS.ReadSkins("显示投票")
	S=Replace(S,"{投票选项}","<form style='margin:0' method='post' action='submit.asp?action=vote&id="&id&"&type="&Arr_Rs(3,0)&"'>"&VoteShow)
	S=Replace(S,"{投票信息}",ClueTxt&"</form>")
	ShowVote=S
End Function

Function ShowBBS()
	Dim arr_Rs,i,P,IUBB,Grade,ToUrl,Temp,Temp1,TopicN
	Dim S,Template,TempStr,Lou,Fontsize,ShowCaption,ShowMood,AppraiseInfo
	ToUrl="boardid="&BBS.boardid&"&ID="&id&"&TB="&BBS.TB
	Set IUBB=New Cls_IUBB
	Set P = New Cls_PageView
	p.strTableName = "[Bbs"&BBS.TB&"] As B inner join [User] As U on B.Name=U.Name"
	p.strPageUrl = "?"&ToUrl
						' 0        1         2        3       4        5        6       7        8           9   10    11       12     13   14     15     16     17         18       19       20   21    22      23      24       25     26       27       28      29      30     31         32         33     34             35
	p.strFieldsList = "B.BbsID,B.TopicID,B.Face,B.Caption,B.Content,B.Name,B.AddTime,B.boardid,B.UbbString,B.IP,U.Id,U.Name,U.IsQQpic,U.QQ,U.Pic,U.Picw,U.Pich,U.GradeID,U.EssayNum,U.Mail,U.Home,U.Sex,U.Mark,U.Coin,U.Sign,U.Regtime,U.IsShow,U.IsDel,U.IsSign,U.IsVip,U.RegIp,U.LoginNum,U.Honor,U.Faction,B.IsAppraise,B.IsDel"
	p.strCondiction = "B.isDel<>1 and (B.topicid="&id&" or B.replytopicid="&id&")"
	p.strOrderList = "B.BbsID"
	p.strPrimaryKey = "BbsID"
	P.CountSQL=TopicReplyNum+1
	p.intPageSize = BBS.Info(80)
	p.intPageNow = Request.QueryString("page")
	p.strCookiesName = "Show_"&ID
	p.Reloadtime=2
	p.InitClass
	Arr_Rs = p.arrRecordInfo
	PageInfo = p.strPageInfo
	page=p.intPageNow
	Set p = nothing
	Template=BBS.ReadSkins("显示帖子")
  If IsArray(Arr_Rs) Then 
	Lou = (Page-1)*10
	For i = 0 to UBound(Arr_Rs, 2)
	UserName=Arr_Rs(11,i)
	BBSID=Arr_Rs(0,i)
	IUBB.UbbString=Arr_Rs(8,i)
	S=Replace(Template,"{数据表ID}",BBS.TB)
	S=Replace(S,"{版块ID}",BBS.boardid)
	S=Replace(S,"{主题ID}",ID)
	S=Replace(S,"{页次}",page)
	S=Replace(S,"{帖子ID}",Arr_Rs(0,i))
	S=Replace(S,"{发帖时间}",Arr_Rs(6,i))
	Temp="*.*.*.*"
	If BBS.Founduser Then
		IF SESSION(CacheName& "MyGradeInfo")(42)="1" then
			Temp=BBS.Fun.GetSqlStr(Arr_Rs(9,i))
		End If
	End If
	S=Replace(S,"{用户IP}",Temp)
	If i mod 2 =0 then Temp=" style=""background:"&BBS.Skinspic(1)&"""" Else Temp=" style=""background:"&BBS.Skinspic(2)&""""
	S=Replace(S,"{背景色}",Temp)	
	S=Replace(S,"{QQ}",BBS.Fun.GetSqlStr(Arr_Rs(13,i)))
	S=Replace(S,"{邮箱}",BBS.Fun.GetSqlStr(Arr_Rs(19,i)))
	S=Replace(S,"{主页}",BBS.Fun.GetSqlStr(Arr_Rs(20,i)))
	S=Replace(S,"{性别}",BBS.Fun.GetSqlStr(Arr_Rs(21,i)))
	S=Replace(S,"{积分数}",BBS.Fun.GetSqlStr(Arr_Rs(22,i)))
	S=Replace(S,"{金钱数}",BBS.Fun.GetSqlStr(Arr_Rs(23,i)))
	S=Replace(S,"{帖数}",Arr_Rs(18,i))
	S=Replace(S,"{注册时间}",formatdatetime(Arr_Rs(25,i),2))
	Grade=BBS.GetGradeInfo(Arr_Rs(17,i))
	Grade=split(Grade,"|")
	S=Replace(S,"{等级图片}","<img src='pic/grade/"&Grade(3)&"' alt='"&Grade(2)&"' />")
	If len(Grade(4))>3 Then Temp="<img src='pic/grade/"&Grade(4)&"' alt='尊贵身份标志' />" Else Temp=""
	S=Replace(S,"{等级标志图片}",Temp)
	S=Replace(S,"{用户名称}","<span style=""color:"&Grade(6)&";"">"&Arr_Rs(11,i)&"</span>")
	S=Replace(S,"{等级}",Grade(2))
	Temp=BBS.Fun.GetSqlStr(Arr_Rs(32,i))
	If Temp="" Then Temp="社区公民" Else Temp="<font color=#2779A5>"&Temp&"</font>"
	S=Replace(S,"{头衔}",Temp)
	Temp=BBS.Fun.GetSqlStr(Arr_Rs(33,i))
	If Temp="" Then Temp="无帮无派"
	S=Replace(S,"{帮派}",Temp)
	S=Replace(S,"{用户}",Arr_Rs(11,i))

	If IsOnline(Arr_Rs(11,i)) Then
		S=Replace(S,"{在线状态}",BBS.SkinsPIC(19))
	Else
		S=Replace(S,"{在线状态}",BBS.SkinsPIC(20))
	End If
	If Arr_Rs(27,i) then
		S=Replace(S,"{用户头像}","<font color=""#F00""><b>系统警告<br />此用户已被暂时删除！</b></font>")
	Else
		IF Arr_Rs(12,i) then
			S=Replace(S,"{用户头像}","<img src='http://qqshow-user.tencent.com/"&Arr_Rs(13,i)&"/10/' alt='"&Arr_Rs(11,i)&"' />")
		else
			S=Replace(S,"{用户头像}","<img src='"&Arr_Rs(14,i)&"' width='"&Arr_Rs(15,i)&"' height='"&Arr_Rs(16,i)&"' alt='"&Arr_Rs(11,i)&"' />")
		End If
	End If
	If Arr_Rs(1,i)=ID Then
		ShowCaption=TopicCaption
		ShowMood=TopicMood()
		S=Replace(S,"{楼号}","楼主")
		S=Replace(S,"{楼数}",0)
		TopicN = 0
	Else
		ShowCaption="":ShowMood=""
		S=Replace(S,"{楼号}","第 <font color=#FF0000>"&Lou+i&"</font> 楼")
		S=Replace(S,"{楼数}",Lou+i)
		TopicN = Lou+i
	End If
	If Arr_Rs(35,i)=2 Then
		Temp1="解除屏蔽"
	Else
		Temp1="单帖屏蔽"
	End If
	Temp=toUrl&"&BbsID="&Arr_Rs(0,i)&"&page="&Page
	S=S&"<div id="""&Arr_Rs(0,i)&""" class=""menu""><div style=""width:80px""><div class=""menuitems""><a href=""post.asp?action=edit&"&Temp&""">编辑帖子</a></div><div class=""menuitems""><a href=""settopic.asp?action=评帖&"&Temp&""">帖子评价</a></div><div class=""menuitems""><a href=""settopic.asp?action=屏蔽&"&Temp&""">"&Temp1&"</a></div><div class=""menuitems""><a href=""settopic.asp?action=删除&"&Temp&""">删除帖子</a></div></div></div>"
	AppraiseInfo=""
	If Arr_Rs(34,i)=1 Then AppraiseInfo=Appraise(Arr_Rs(0,i))
    Temp="<div style=""height:auto!important;height:300px;min-height:300px;line-height:normal;margin-top:10px;word-wrap:break-word;word-break:break-all"">"		
    Temp=Temp&""&ShowMood&"<div style=""margin-bottom:8px;padding-bottom:5px;font-size:14px;color:#0000ff"">"&ShowCaption&"</div>"
	IF Arr_Rs(26,i)=1 then
		S=Replace(S,"{签名}","<div class='cover'>签名档已被管理员屏蔽</div>")
		S=Replace(S,"{帖子内容}",Temp&"<div class='cover'>该用户帖子已被管理员屏蔽，请和管理员联系</div>")
	Else
		Temp1=Arr_Rs(24,i)
		IF BBS.Info(44)="0" then Temp1="" 
		IF Arr_Rs(28,i)=1 Then Temp1="<div class='cover'>签名档已被管理员屏蔽</div>"
		IF isNull(Temp1) or Temp1="" then
			S=Replace(S,"{签名}","<font color=#999999>这家伙很懒，什么也没有留下！</font>")
		Else
			S=Replace(S,"{签名}",IUBB.Sign_Code(Temp1))
		End IF
		If Arr_Rs(35,i)=2 Then
			S=Replace(S,"{帖子内容}",Temp&"<div class='cover'>内容被屏蔽</div></div>")
		Else
			S=Replace(S,"{帖子内容}",Temp&"<div id=""textstyle_"&TopicN&""" style=""font-size:14px"">"&IUBB.UBB(Arr_Rs(4,i),1)&"</div></div>"&AppraiseInfo)
		End If
	End IF
	TempStr=TempStr&S
	Next
	ShowBBS=TempStr
  End If
  Set IUBB=Nothing
End Function

		
Function IsOnline(str)
	dim EachOnline,S,i,OnlineCache
	IsOnline=False
	OnlineCache=BBS.Cache.Value("OnlineCache")
	If Instr(lcase(OnlineCache),"|"&Lcase(str)&"|")<>0 Then
	IsOnline=True
	End If
End Function


Function SetTopic()
	Dim S,GO
	If Not BBS.FoundUser Then Exit Function
		GO="<a href='settopic.asp?boardid="&BBS.boardid&"&ID="&id&"&TB="&BBS.TB&"&Action=GO'>GO</a> | "
		S="主题管理："
		If BBS.MyAdmin >= 7 Then
		  IF TopicTopType=5 then S=S&Replace(GO,"GO","取消总置顶") Else S=S&Replace(GO,"GO","总置顶")
		  IF TopicTopType=4 Then S=S&Replace(GO,"GO","取消区置顶") ELse S=S&Replace(GO,"GO","区置顶")
		  IF TopicTopType=3 Then S=S&Replace(GO,"GO","取消置顶") ELse S=S&Replace(GO,"GO","置顶")
		  IF TopicIsGood=1 Then S=S&Replace(GO,"GO","取消精华") Else S=S&Replace(GO,"GO","精华")
		  IF TopicIsLock=1 Then S=S&Replace(GO,"GO","解锁") Else S=S&Replace(GO,"GO","锁定")
		  S=S&Replace(GO,"GO","删除")&Replace(GO,"GO","移动")&Replace(GO,"GO","提升")&Replace(GO,"GO","沉底")
		End If
		S=S&Replace(Replace(GO,"GO","已解决")," | ","")
		SetTopic=S
End Function

Function Appraise(AstID)
	Dim Rs,Arr_Rs,i
	Set Rs=BBS.Execute("Select BbsID,Cause,Mark,Coin,GameCoin,Adminname,AddTime From [Appraise] where BBSID="&AstID&" And TopicID="&id&" order by AddTime desc")
	If Rs.Eof Then
		Exit Function
	Else
		Arr_Rs=Rs.GetRows(-1)
		Rs.Close
		Appraise="<div class='appraise'>此帖被评"
		If BBS.FoundUser Then
		If SESSION(CacheName& "MyGradeInfo")(41)="1" Then Appraise=Appraise&"【<a href=settopic.asp?action=delappraise&TB="&BBS.TB&"&ID="&id&"&BbsID="&AstID&"&boardid="&BBS.boardid&">删除</a>】"
		End If
		For i=0 To Ubound(Arr_Rs,2)
		Appraise=Appraise&"<br /><span style='color:#00F'>"&Arr_Rs(1,i)&"</span> "
		If Arr_Rs(2,i)<>0 Then Appraise=Appraise&BBS.Info(121)&":<span style='color:#F00'>"&Arr_rs(2,i)&"</span> "
		If Arr_Rs(3,i)<>0 Then Appraise=Appraise&BBS.Info(120)&":<span style='color:#F00'>"&Arr_rs(3,i)&"</span> "
		If Arr_Rs(4,i)<>0 Then Appraise=Appraise&BBS.Info(122)&":<span style='color:#F00'>"&Arr_rs(4,i)&"</span> "
		Appraise=Appraise&"<span style='color:#AAA'>"&Arr_rs(5,i)&" "&Arr_rs(6,i)&"</span><br />"
		Next
		Appraise=Appraise&"</div>"
	End If
End Function

Function FastReply()
	with BBS
	If Not .FoundUser Then Exit Function
	If .BoardType<>2 or .MyAdmin=9 or .MyAdmin=8 or (.MyAdmin=7 And .IsBoardAdmin) Then
	IF TopicIsLock=0 Or .MyAdmin=9 or .MyAdmin=8 then
		Dim Tmp,S,Edit
		S="<form style='margin:0;' name='preview' action='preview.asp?action=preview' method='post' target='preview'><input type='hidden' name='pcaption' /><input type='hidden' name='pcontent' /></form>"
		S=S&"<form style='margin:0;' method=POST name='say' action='postsave.asp?action=reply&boardid="&.boardid&"&TB="&.TB&"&ID="&id&"&page=100' ><input type=hidden name='caption' id='caption' value='回复:"&TopicCaption&"' /><input id='content' name='content' type='hidden' value='' />"
			Tmp="<br />&nbsp;HTML标签：<font color=red>"
			If .Info(60)="1" Then Tmp=Tmp&"×" Else Tmp=Tmp&"√"
			Tmp=Tmp&"</font><br />&nbsp;UBB 标签：<font color=red>√</font><br />上传文件：<font color=red>"
			If .Info(30)="0" Then Tmp=Tmp&"×" Else Tmp=Tmp&"√"
			Tmp=Tmp&"</font><br />内容最多：<font color=red>30KB</font><br />"
		If .Info(60)="1" Then Edit="UbbEdit()" Else Edit="HtmlEdit()"
		Edit="<script type=""text/javascript"">"&Edit&"</script>"
		IF .Info(15)="1" then Edit=Edit&"验证码："&.GetiCode Else Edit=Edit&"<input type=hidden name='iCode' id='iCode' value='BBS' />"
		S=S&.Row("<a href='post.asp?action=reply&boardid="&.boardid&"&id="&id&"'>高级回复</a><br /><b>帖子内容：</b>"&Tmp,Edit,"80%","")
	S=S&"<div align='center' style=""padding:5px""><input type='button' value='OK 发表' id='sayb' onclick='checkform()' class='button' /> <input type=button value='预 览' onclick='Gopreview()' class='button' /> <input type='reset' value='NO 重写' onclick='Goreset()' class='button' /></div></form>" 
	.ShowTable "快速回复:"&TopicCaption,S
	End If
	End If
	End With
End Function
%>