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
	BBS.Position=BBS.Position&" -> <a href='userinfo.asp'>用户控制面版<a>"
	BBS.Head "help.asp"&Page_Url,"","查看等级权限"
Else
	BBS.Head "help.asp"&Page_Url,"","查看论坛帮助"
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
Content="<div align='center'><b>====== 论坛帮助目录 ======</b><table border='0' cellspacing='10' cellpadding='0'><tr><td><a href=?action=what>简约论坛的常规说明</a></td><td><a href=?action=grade>关于用户等级的帮助</a></td><td><a href=?action=forget>关于遗忘密码的帮助</a></td><td><a href=?action=upload>关于文件上传的帮助</a></td></tr><tr><td><a href=?action=sms>关于论坛信箱的帮助</a></td><td><a href=?action=usersetup>关于更改信息的帮助</a></td><td><a href=?action=Say>关于发表帖子的帮助</a></td><td><a href=?action=ubb>关于UBB功能的帮助</a></td></tr></table></div>"
BBS.ShowTable "论坛帮助",Content
End Sub



Sub GradeString()
Dim Rs,ID,GradeName,Pic,Spic,Grouping,EssayNum,Title
Dim S,GS,T,Y,N
ID=BBS.Checknum(request.querystring("ID"))
Title="等级权限"
If ID=0 Then
	If Not BBS.FoundUser Then
		BBS.GoToerr(26)
	Else
		Title="我的等级权限"
		Response.Write BBS.ReadSkins("用户控制面版")
		ID=SESSION(CacheName & "MyInfo")(15)
	End If
End If
Y="<span style='color:#F00'>√</span>"
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
S=BBS.Row("等级名称:","<b>"&GradeName&"</b>","65%","")
If Grouping=0 Then S=S&BBS.Row("必需达到帖数：",EssayNum&" 篇","65%","")
If len(Pic)>3 Then T="<img src='Pic/Grade/"&pic&"' />" Else T="无"
S=S&BBS.Row("等级图片：",T,"65%","18px")
If len(SPic)>3 Then T="<img src='Pic/Grade/"&spic&"' />" Else T="无"
S=S&BBS.Row("身份标志图片",T,"65%","18px")
S=S&"<div class='title'>基本权限</div>"
S=S&BBS.Row("帖子显示名字颜色","<div style=""width:20px;height:20px;BACKGROUND:"&Gs(0)&""">&nbsp;</div>","65%","")
If Gs(1)="1" Then T=Y Else T="×"
S=S&BBS.Row("是否可以修改自己资料：",T,"65%","")
If Gs(2)="1" Then T=Y Else T="×"
S=S&BBS.Row("是否可以自定义头衔：",T,"65%","")
S=S&BBS.Row("发帖子最大的字符数：","<font color='#FF0000'>"&Gs(3)&"</font> 字节","65%","")
If Gs(4)="1" Then T=Y Else T="×"
S=S&BBS.Row("是否可以发表醒目标题：",T,"65%","")
If Gs(5)="1" Then T=Y Else T="×"
S=S&BBS.Row("是否可以参加投票活动：",T,"65%","")
If Gs(6)="1" Then T=Y Else T="×"
S=S&BBS.Row("是否可以发表投票主题：",T,"65%","")
If Gs(8)="1" Then T=Y Else T="×"
S=S&BBS.Row("是否可以上传：",T,"65%","")
S=S&BBS.Row("一天的上传个数：","<font color='#FF0000'>"&GS(9)&"</font> 个","65%","")
S=S&BBS.Row("每个上传大小：","<font color='#FF0000'>"&GS(10)&"</font> KB","65%","")
If Gs(11)="1" Then T=Y Else T="×"
S=S&BBS.Row("是否可以上传头像：",T,"65%","")
S=S&BBS.Row("论坛信箱最大条数：","<font color='#FF0000'>"&GS(12)&"</font> 条","65%","")
S=S&BBS.Row("限制每天发送信件的次数：","<font color='#FF0000'>"&GS(7)&"</font> 次","65%","")
S=S&BBS.Row("限制每封信字符数","<font color='#FF0000'>"&GS(13)&"</font> 字节","65%","")
If Gs(14)="1" Then T=Y Else T="×"
S=S&BBS.Row("是否可以搜索论坛：",T,"65%","")
If Gs(15)="1" Then T=Y Else T="×"
S=S&BBS.Row("是否可以查看他人信息：",T,"65%","")
If Gs(16)="1" Then T=Y Else T="×"
S=S&BBS.Row("可以不受时间限制编辑自己帖子：",T,"65%","")
If Gs(17)="1" Then T=Y Else T="×"
S=S&BBS.Row("开启可以删除自己的帖子：",T,"65%","")
S=S&"<div class='title'>管理权限</div>"
If Gs(18)="1" Then T=Y Else T="×"
S=S&BBS.Row("可以编辑帖子：",T,"65%","")
If Gs(19)="1" Then T=Y Else T="×"
S=S&BBS.Row("编辑帖子可以不留下蛛迹：",T,"65%","")
If Gs(20)="1" Then T=Y Else T="×"
S=S&BBS.Row("可以删除帖子：",T,"65%","")
If Gs(21)="1" Then T=Y Else T="×"
S=S&BBS.Row("可以屏蔽帖子：",T,"65%","")
If Gs(22)="1" Then T=Y Else T="×"
S=S&BBS.Row("可以移动帖子：",T,"65%","")
If Gs(23)="1" Then T=Y Else T="×"
S=S&BBS.Row("可以提升主题：",T,"65%","")
If Gs(24)="1" Then T=Y Else T="×"
S=S&BBS.Row("可以沉底主题：",T,"65%","")
If Gs(25)="1" Then T=Y Else T="×"
S=S&BBS.Row("可以(设/解)置顶主题：",T,"65%","")
If Gs(26)="1" Then T=Y Else T="×"
S=S&BBS.Row("可以(设/解)区置顶主题：",T,"65%","")
If Gs(27)="1" Then T=Y Else T="×"
S=S&BBS.Row("可以(设/解)总置顶主题：",T,"65%","")
If Gs(28)="1" Then T=Y Else T="×"
S=S&BBS.Row("可以(设/解)精华主题：",T,"65%","")
If Gs(29)="1" Then T=Y Else T="×"
S=S&BBS.Row("可以(设/解)锁定主题：",T,"65%","")
If Gs(30)="1" Then T=Y Else T="×"
S=S&BBS.Row("可以进行评帖奖罚操作：",T,"65%","")
If Gs(31)="1" Then T=Y Else T="×"
S=S&BBS.Row("可以不需要投票可查投票详情：",T,"65%","")
If Gs(32)="1" Then T=Y Else T="×"
S=S&BBS.Row("可以编辑投票的选项：",T,"65%","")
If Gs(33)="1" Then T=Y Else T="×"
S=S&BBS.Row("可以不受特殊帖限制：",T,"65%","")
If Gs(34)="1" Then T=Y Else T="×"
S=S&BBS.Row("可以发布论坛公告：",T,"65%","")
If Gs(35)="1" Then T=Y Else T="×"
S=S&BBS.Row("可以删除评帖记录：",T,"65%","")
S=S&"<div class='title'>高级管理权限</div>"
If Gs(36)="1" Then T=Y Else T="×"
S=S&BBS.Row("可以查看用户IP：",T,"65%","")
If Gs(37)="1" Then T=Y Else T="×"
S=S&BBS.Row("可以查看论坛日志：",T,"65%","")
'If Gs(38)="1" Then T=Y Else T="×"
'S=S&BBS.Row("开启可以批量操作主题：",T,"65%","")
If lcase(Action)<>"mygrade" Then
	S=S&"<div style="" padding:5px;BACKGROUND: "&BBS.SkinsPIC(2)&";"" align=""center""><a href='javascript:history.go(-1)'>【返回】</a><a href=help.asp>【返回帮助目录】</a></div></form>"
End If
BBS.ShowTable Title,S
End Sub


Sub Grade()
	Dim ARs,i,S,AdminGrade,EssayGrade,otherGrade,EssayNum,Temp
	ARs=BBS.SetGradeInfoCache()
	For i=0 To Ubound(ARs,2)
	If ARs(1,i)=0 Then EssayNum="必需达到帖数：<span style='color:#F00'>"&ARs(5,i)&"</span>" Else EssayNum=""
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
		S="<div class='title'>系统等级组</div>"&AdminGrade&"<div class='title'>特殊等级组</div>"&otherGrade&"<div class='title'>常规等级组</div>"&EssayGrade
		S=S&"<div style=""padding:3px;BACKGROUND: "&BBS.SkinsPIC(2)&";"" align=""center""><a href='help.asp'>【返回帮助目录】</a></div>"
	BBS.ShowTable"论坛等级",S
End sub

Sub Upload()
	Content="<br><br><div align='center'><b>====== 论坛上传帮助 ======</b></div><blockquote><li>发表帖子时，点击功能按纽上的“浏览”找到要上传的文件，选定后，然后点击“上传”按纽即可上传。<li>每日的上传大小和次数限制，本论坛根据每个会员等级做不同的规定。<li>本论坛允许上传的文件："&Replace(BBS.Info(34)&"|"&BBS.Info(35),"|","、")&"<div align='center'><a href=help.asp>【返回帮助目录】</a></div></blockquote>"
	BBS.ShowTable"论坛帮助",Content
End Sub

Sub Sms()
	Content="<br><br><div align='center'><b>====== 论坛信箱帮助 ======</b></div><blockquote>论坛信箱功能也相当于留言，所不同的是他可以在用户登录社区的情况下即时收发消息，方便简洁。<ul><li><b>发送消息</b>：登录用户方能发送消息，一种是自己填写收件人名称（该用户必须是论坛注册用户）和完整的标题和内容，内容支持ubb格式；另外一种是在论坛中查看帖子的时候直接给作者发送消息，需要填写内容同后点“发送”即可<li><b>收件箱</b>：登录论坛后点击上方“用户助手”下的“留言板”，列出所有已读和未读的消息及其标题、发件人，可以进行读取信息和全部删除操作。<li><b>新消息</b>：登录论坛后，每当有别人给你发送新的短消息，或者原来的消息还未读取，论坛都将有提示，直接点击后阅读。</ul><div align='center'><a href=help.asp>【返回帮助目录】</a></div></blockquote>"
	BBS.ShowTable"论坛帮助",Content
End sub

Sub Forget()
	Content="<br><br><div align='center'><b>====== 关于遗忘密码的帮助 ======</b></div><blockquote><ul><li>如果遗忘密码，你可以用注册时填写的密码问题和密码答案来<a href=UserSetup.asp?action=ForgetPassword>取回密码</a>。</ul><ul><li>如果你忘记了密码提示和密码答案，请与论坛管理员联系，由他（她）来为你设定新的密码。</ul><div align='center'><a href=help.asp>【返回帮助目录】</a></div></blockquote>"
	BBS.ShowTable"论坛帮助",Content
End sub

Sub What()
Content="<br><br><div align='center'><b>====== 常规帮助 ======</b></div>"&_
"<ul><li><b>怎样才能加入论坛？</b></li>"&_
"<ul><li>您可以点击论坛顶部的“会员注册”注册为本站会员，请将必填资料认真填写，E-mai地址输入正确有效，以便您能正常注册和使用密码找回功能。您当然可以不必注册为我们的会员，但是为了您能够使用本论坛的全部功能，我们仍建议您注册。</li></ul></ul>"&_
"<ul><li><b>"&BBS.Info(121)&"和"&BBS.Info(122)&"有什么用? 我的金钱高和积分高有什么好处？</b></li>"&_
"<ul><li>我们用金钱值和积分点数来活跃论坛的气氛。</li>"&_
"<li>积分代表在论坛身份的尊贵。一般拥有较多的积分的会员代表对论坛的贡献程度较多。</li>"&_
"<li>金钱代表在论坛的虚拟货币，拥有较多的金钱的会员可以玩论坛更多的一些误乐插件。</li>"&_
"<li>有的帖子是需要有一定金钱和积分才可以浏览的，这也是对某些发帖的网友的尊重，请大家互相体谅。</li>"&_
"<li>金钱值只是说明网友在本论坛的活跃情况，并不一定代表网友任何方面的个人水平。</li>"&_
"<li>我们不会根据网友积分情况的不同，而对网友本人有任何方式的优待或歧视。</li>"&_
"<li>请网友正确对待金钱积分这种评定方式，不要随意灌水以骗取金钱积分如果恶意灌水，将被视为对社区的恶意攻击。</li>"&_
"<li>我们将对恶意灌水的网友进行处罚，删除帐号，并保留进一步采取行动的权利。</li></li></ul></ul>"&_
"<ul><li><b>分值是如何计算的？</b></li>"&_
"<ul>"&_
"<li>发表一篇文章"&BBS.Info(120)&"增"&BBS.Info(102)&"，"&BBS.Info(121)&"增"&BBS.Info(103)&"，"&BBS.Info(122)&"增"&BBS.Info(104)&"，回复帖子增加金钱30(同时增加该主题帖作者相同的金钱30)</li>"&_
"<li>帖子被管理员或者斑竹设为精华后，发帖人发贴人"&BBS.Info(120)&"增"&BBS.Info(99)&"，"&BBS.Info(121)&"增"&BBS.Info(100)&"，"&BBS.Info(122)&"增"&BBS.Info(101)&"，取消精华后，则相应减少！</li>"&_
"<li>帖子被管理员或者斑竹设为置顶后，发贴人发贴人"&BBS.Info(120)&"增"&BBS.Info(96)&"，"&BBS.Info(121)&"增"&BBS.Info(97)&"，"&BBS.Info(122)&"增"&BBS.Info(98)&"，取消置顶后，则相应减少！</li>"&_
"<li>帖子被管理员设为区置顶后，发贴人"&BBS.Info(120)&"增"&BBS.Info(93)&"，"&BBS.Info(121)&"增"&BBS.Info(94)&"，"&BBS.Info(122)&"增"&BBS.Info(95)&"，取消区置顶后，则相应减少！</li>"&_
"<li>帖子被管理员总置顶后，发贴人"&BBS.Info(120)&"增"&BBS.Info(90)&"，"&BBS.Info(121)&"增"&BBS.Info(91)&"，"&BBS.Info(122)&"增"&BBS.Info(92)&"，取消总置顶后，则相应减少！</li></ul></ul>"&_
"<ul><li><b>怎样才能知道自己的积分和积分的排名情况？</b></li>"&_
"<ul><li>只要在论坛里找到自己的用户名，点击即可查看自己的积分。</li>"&_
"<li>对于论坛排名，您可以通过点论坛菜单的<a href=userboard.asp>用户列表</a>查看。</li></ul></ul>"&_
"<ul><li><b>怎样才能在帖子后加上签名？ </b></li>"&_
"<ul><li>您可以通过论坛顶部“用户助手”下“修改资料”个性签名一栏中写入您的个性签名。</li>"&_
"<li>签名支持UBB，可以使用图片，格式是[img]图片地址[/img]。</li></ul></ul>"&_
"<ul><li><b>如何快速找到需要的文章？</b></li>"&_
"<ul><li>您可以使用论坛的搜索功能，搜索整个论坛您可以在“论坛菜单”下的“<a href=Search.asp>论坛搜索</a>”内写入满足的条件即可搜索。另外在各版面下方也可以按照帖子内容搜索。还可以选择使用“今日新帖”、“一周内新帖” “上次到访后的新帖”、“最旺人气帖”、“最旺回复帖”等进行简易搜索。</li></ul></ul>"&_
"<ul><li><b>什么是精华区？帖子是谁将它加入精华区的？如何查看？</b>"&_
"<ul><li>精华区是论坛版面存放相对有价值，技术含量较高或内容比较有意义的帖子的，网友通常可以在精华区内找到很多有用的东西。本论坛的每个版面都有自己的精华区。精华区由版主管理。版主可以将版面上的帖子加入到精华区。并可编辑进行再加工。</li>"&_
"<li>精华区的帖子即使在空间容量不够的情况下也不会被删除，会被永久保留！</li>"&_
"<li>只要进入相关版面，在版面的右上方就可以看到”本版精华“字样，点击即可查看；在“论坛菜单”下也有一个“精华区”，那是全部论坛的精华所在，不建议您以此种方式查看，如果您想找相关精华请到相关版面查找。</li></ul></ul>"&_
"<ul><li><b>成为版主的条件是什么？有哪些权利？</b></li>"&_
"<ul><li>版主当然有所见长之处，并且为人热心，愿意为网友无偿服务。</li>"&_
"<li>拥有一定群众基础，能花时间维护论坛的。</li>"&_
"<li>版主可以查询版面的帖子，可以任意删除或编辑版面的帖子。</li>"&_
"<li>版主可以把帖子加锁（解锁），置顶（解除置顶），提升和沉底帖子，加入精华区（解除精华），发布版面公告。</li>"&_ 
"<li>版主可以在所属版面上发布或删除公告。</li>"&_
"<li>如果论坛开启了版主继承功能，上级版主还可以管理下级论坛的帖子。</li></ul></ul><div align='center'><a href=help.asp>【返回帮助目录】</a></div></blockquote>"
BBS.ShowTable"论坛帮助",Content
End sub

Sub UserSetup()
Content="<br><br><div align='center'><b>====== 更改个人信息帮助 ======</b></div><blockquote>"&_
"只有登录用户才能进行此项操作，原注册用户名不可修改，请在论坛顶部“助手”下找到“修改资料”进入，可以更新的信息如下："&_
"<li>密码：登录论坛所用"&_
"<li>Email地址：必须填入正确合法的邮箱地址"&_
"<li>生日：在你的生日那天在论坛会出现你生日的提示（面向所有会员）"&_
"<li>个人主页：可选项（有的话建议填上，让大家见识下）"&_
"<li>OICQ：可选项（为方便联系，建议填上）"&_
"<li>个性头像：所选择头像将在帖子中出现，可自行连接图片url或QQ形象作为头像"&_
"<li>个性签名：支持ubb，如果填入，将出现在文章的结尾<br><br><div align='center'><a href=help.asp>【返回帮助目录】</a></div></blockquote>"
BBS.ShowTable"论坛帮助",Content
End sub

Sub Say()
Content="<br><br><div align='center'><b>====== 发表帖子帮助 ======</b></div><blockquote>"&_
"<ul><li>只有注册并且已登陆的用户，才可以发起一个新主题，或是回复已有主题。"&_
"<li>发起特殊帖的帮助说明:</B>"&_
"<ul><li>【文字大小】选择你要的字号，在出现的输入栏填上你的内容即可。"&_
"<li>【文字颜色】选择你要的颜色，在出现的输入栏填上你的内容即可。"&_
"<li>【回复可见】该帖子内容只有回复了该主题的用户才可看见。<br>在发帖时点“特殊帖”的“回复可见”，这时会在帖子框内出现<font color=blue>[reply]内容[/reply]</font>，将其中的“内容”替换为你自己的内容即可。"&_
"<li>【指定读者】该内容只有被指定的注册用户才可见。<br>在发帖时点“特殊帖”“指定读者”，这时会在帖子框内出现<font color=blue>[UserName=admin]内容[/UserName]</font>标签，将其中的“admin”替换为某个注册用户，将“内容”替换为你的内容即可。"&_
"<li>【金钱可见】该帖子内容只有达到指定金钱值的用户才可看。<br>点击“金钱可见”这时会在帖子框内出现<font color=blue>[COIN=1000]内容[/COIN]</font>将其中的1000替换为你想要的金钱数，将“内容”替换为你想要的内容即可。"&_
"<li>【积分可见】该帖子内容只有达到指定积分值的用户才可看。<br>点击点击“积分可见”这时会在帖子框内出现<font color=blue>[MAKE=3]内容[/MAKE]</font>将其中的3替换为你想要的金钱数，将“内容”替换为你想要的内容即可。"&_
"<li>【付费可见】该功能可以设置“实用价值高”的帖子的阅读价格，阅读者需购买才能看。<br>在发帖时点“特殊帖”的“付费可见”，这时会现在一个输入栏，填上帖子的价格，这时会在帖子框内出现<font color=blue>[BUYPOST=100]内容[/BUYPOST]</font>标签，将“内容”替换为你自己的内容即可。"&_
"<li>【日期可见】该帖子内容只有到了规定日期后才可以看见。<br>在发帖时点“特殊帖”的“日期可见”，这时会现在一个输入栏，填上帖子的可见日期，这时会在帖子框内出现<font color=blue>[DATE=2003-10-1]内容[/DATE]</font>标签，将“内容”替换为你自己的内容即可。"&_
"<li>【性别可见】该帖子内容只有指定的性别才可以看见。<br>在发帖时点“特殊帖”的“性别可见”，这时会现在一个输入栏，填上1或0（1代表男，0代表女），这时会在帖子框内出现<font color=blue>[SEX=1]内容[/SEX]</font>标签，将“内容”替换为你自己的内容即可。"&_
"<li>【登陆可见】该内容只有登陆的用户才可见。<br>在发帖时点“特殊帖”“登陆可见”，这时会在帖子框内出现<font color=blue>[LOGIN]内容[/LOGIN]</font>标签，将“内容”替换为你的内容即可。</ul>"&_
"<li>每张帖子下面都有快速回复栏，你可以直接在里面输入内容回复。"&_
"<li>如果想要用到某些标签如发特殊帖、上传文件等，请点击帖子左上方的“回复帖子”即可，回复帖子时与发帖相同."&_
"<li>除了回复帖子时不用主题描述项外（您可以根据你所发帖子的类型选择相应的“主题标志”，你要确保填写所有项以便成功发出帖子。"&_
"<li>点击“功能按纽”一栏的按纽将会帮助你快速插入某些UBB标签。（<a href=?action=Ubb>点击进入UBB帮助</a>）</ul></ul><div align='center'><a href=help.asp>【返回帮助目录】</a></div>"
BBS.ShowTable"论坛帮助",Content
End Sub

Sub Ubb()
Content="<br><br><div align='center'><b>====== UBB标签帮助 ======</b></div><blockquote>"&_
"<ul>UBB标签就是不允许使用HTML语法的情况下，通过论坛的特殊转换程序，以至可以支持少量常用的、无危害性的HTML效果显示。以下为具体使用说明："&_
"<p><font color=red>[B]</font><b>文字</b><font color=red>[/B]</font><br>在文字的位置可以任意加入您需要的字符，显示为粗体效果。"&_
"<p><font color=red>[I]</font><i>文字</i><font color=red>[/I]</font><br>在文字的位置可以任意加入您需要的字符，显示为斜体效果。"&_
"<p><font color=red>[U]</font><u>文字</u><font color=red>[/U]</font><br>在文字的位置可以任意加入您需要的字符，显示为下划线效果。"&_
"<p><font color=red>[align=center]</font>文字<font color=red>[/align]</font><br>在文字的位置可以任意加入您需要的字符，center位置center表示居中，left表示居左，right表示居右。"&_
"<p><A HREF='http://www.74177.com/bbs'><font color=red>http://www.74177.com/bbs</font></A><br>直接输入网址，论坛会自动识别"&_
"<P><font color=red>[URL=http://www.74177.com/bbs]</font><A HREF=http://www.74177.com/bbs>简约论坛</A><font color=red>[/URL]</font>：<br>或则你也可以连接具体地址或者文字连接。"&_
"<P><font color=red>[EMAIL]</font><A HREF=""mailto:abc@abc.com"">abc@abc.com</A><font color=red>[/EMAIL]</font><br>"&_
"<font color=red>[EMAIL=MAILTO:abc@abc.com]</font><A HREF=""mailto:abc@abc.com"">信箱</A><font color=red>[/EMAIL]</font>：<br>有两种方法可以加入邮件连接，可以连接具体地址或者文字连接。"&_
"<P><font color=red>[img]</font>http://www.74177.com/bbs/images/pic.gif<font color=red>[/img]</font><br>在标签的中间插入图片地址可以实现插图效果。"&_
"<P><font color=red>[flash]</font>Flash连接地址<font color=red>[/Flash]</font><br>在标签的中间插入Flash图片地址可以实现插入Flash。"&_
"<P><font color=red>[Code]</font>文字<font color=red>[/Code]</font><br>在标签中写入文字可实现html中编号效果。"&_
"<P><font color=red>[quote]</font>引用<font color=red>[/quote]</font><br>在标签的中间插入文字可以实现HTMl中引用文字效果。"&_
"<P><font color=red>[list]</font>文字<font color=red>[/list]</font> <font color=red>[list=a]</font>文字<font color=red>[/list]</font>  <font color=red>[list=1]</font>文字<font color=red>[/list]</font>：<br>更改list属性标签，实现HTML目录效果。"&_
"<P><font color=red>[fly]</font>文字<font color=red>[/fly]</font><br>在标签的中间插入文字可以实现文字飞翔效果，类似跑马灯。"&_
"<P><font color=red>[move]</font>文字<font color=red>[/move]</font><br>在标签的中间插入文字可以实现文字移动效果，为来回飘动。"&_
"<P><font color=red>[light]</font>文字<font color=red>[/light]</font><br>在标签的中间插入文字可以实现文字五颜六色的闪光特效。"&_
"<P><font color=red>[shadow=255,red,2]</font>文字<font color=red>[/shadow]</font><br>在标签的中间插入文字可以实现文字阴影特效，shadow内属性依次为宽度、颜色和边界大小。"&_
"<P><font color=red>[color=颜色代码]</font>文字<font color=red>[/color]</font><br>输入您的颜色代码，在标签的中间插入文字可以实现文字颜色改变。"&_
"<P><font color=red>[size=数字]</font>文字<font color=red>[/size]</font><br>输入您的字体大小，在标签的中间插入文字可以实现文字大小改变。"&_
"<P><font color=red>[face=字体]</font>文字<font color=red>[/face]</font><br>输入您需要的字体，在标签的中间插入文字可以实现文字字体转换。"&_
"<P><font color=red>[em1]</font><br>论坛心情图片代码。其中的数字1到180之间是图片代码。"&_
"<P><div align='center'><a href=help.asp>【返回帮助目录】</a></div></blockquote>"
BBS.ShowTable"论坛帮助",Content
End sub

%>