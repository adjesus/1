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
	.Write"<a name='inc'></a><div class='mian'><div class='top'>论坛系统设置</div><div class='divth'>【<a href='#inc1'>论坛基本信息</a>】【<a href='#inc2'>论坛显示设置</a>】【<a href='#inc3'>上传设置</a>】 【<a href='#inc4'>用户选项</a>】 【<a href='#inc5'>帖子设置</a>】 【<a href='#inc6'>论坛资源分配</a>】</div></div>"
'1-10基本信息
	.Write GoForm("BbsInfo")&"<a name='inc1'></a><div class='mian'><div class='top'>论坛基本信息</div>"
	DIVTR"关闭论坛：","维护期间可设置关闭论坛",GetRadio("info3","开启",Info(3),0)&GetRadio("info3","关闭",Info(3),1),40,1
	DIVTR"关闭论坛显示信息","设置关闭论坛后显示的信息,支持Html语法","<textarea rows='3' name='info4' cols='70'>"&Info(4)&"</textarea>",55,2
	DIVTR"论坛名称","你的论坛名称","<input type='text' class='text' name='info0' size='40' value='"&Info(0)&"' /> ",40,1
	DIVTR"论坛地址：","论坛的访问地址","<input type='text' class='text' name='info1' size='40' value='"&Info(1)&"' />尾部不要加“/”",40,2
	DIVTR"主页地址：","主页的访问地址,如果没有可不填","<input type='text' class='text' name='info2' size='40' value='"&Info(2)&"' />",40,1
	DIVTR"建站日期：","论坛落成开业的日期","<input type='text' class='text' name='info5'  value='"&Info(5)&"' /> (格式：YYYY-M-D)",40,2
	DIVTR"论坛顶部广告：","支持Html语法,风格模板显示代码为{广告}","<textarea rows='3' name='info6' cols='70'>"&Info(6)&"</textarea>",58,1
	DIVTR"论坛版权信息：","论坛底部信息,支持Html语法","<textarea rows='3' name='info7' cols='70'>"&Info(7)&"</textarea>",55,2
	DIVTR"在线人数超时：","设定在线人数的在线时间","<input type='text' class='text' name='info8' size='5' value='"&Info(8)&"' />分钟",40,1
	DIVTR"注册间隔：","同一来源的注册间隔时间,如果不想使用这项功能, 请设置为0","<input type='text' class='text' name='info9' size='5' value='"&Info(9)&"' />分钟",55,2
	DIVTR"登陆间隔：","同一来源的登陆间隔时间,如果不想使用这项功能, 请设置为0","<input type='text' class='text' name='info10' size='5' value='"&Info(10)&"' />分钟",55,1
	DIVTR"发帖间隔：","同一来源的发帖间隔时间,如果不想使用这项功能, 请设置为0","<input type='text' class='text' name='info11' size='5' value='"&Info(11)&"' />秒",55,2
	DIVTR"编辑时间：","普通会员修改自己帖子有效时间，如果不想使用这项功能, 请设置为0","<input type='text' class='text' name='info12' size='5' value='"&Info(12)&"' />分钟",55,1
	DIVTR"搜索间隔时间：","限制每次搜索的时间间隔,管理员不受此限","<input type='text' class='text' name='info17' size='5' value='"&Info(17)&"' /> 秒",40,2
	DIVTR"注册验证码：","",GetRadio("info13","否",Info(13),0)&GetRadio("info13","是",Info(13),1),30,1
	DIVTR"登陆验证码：","",GetRadio("info14","否",Info(14),0)&GetRadio("info14","是",Info(14),1),30,2
	DIVTR"发帖验证码：","",GetRadio("info15","否",Info(15),0)&GetRadio("info15","是",Info(15),1),30,1
	DIVTR"删除论坛日志：","",GetRadio("info16","手工删除",Info(16),0)&GetRadio("info16","自动删除7天前的记录",Info(16),1),30,2
'20-29显示设置
	.Write"</div><a name='inc2'></a><div class='mian'><div class='top'>论坛显示设置<a href='#inc'>▲</a></div>"
	DIVTR"显示系统信息：","包括首页公告、快速登陆",GetRadio("info20","否",Info(20),0)&GetRadio("info20","是",Info(20),1),40,1
	DIVTR"显示密码：","请选择否-无需更改",GetRadio("info21","否",Info(21),0),40,2
	DIVTR"显示会员生日：","是否显示首页的会员生日信息",GetRadio("info22","否",Info(22),0)&GetRadio("info22","是",Info(22),1),40,1
	DIVTR"显示论坛联盟：","是否显示论坛首页的友情连接",GetRadio("info23","否",Info(23),0)&GetRadio("info23","是",Info(23),1),40,2
	DIVTR"显示数据查询：","是否显示论坛底部的数据查询",GetRadio("info24","否",Info(24),0)&GetRadio("info24","是",Info(24),1),40,1
	DIVTR"显示执行时间：","是否显示页面下部的加载时间",GetRadio("info25","不显示",Info(25),0)&GetRadio("info25","以毫秒显示",Info(25),1)&GetRadio("info25","以秒显示",Info(25),2),40,2
	DIVTR"显示计数器：","设置论坛的访问计数器",GetRadio("info26","失效",Info(26),0)&GetRadio("info26","显示",Info(26),1)&GetRadio("info26","不显示",Info(26),2),40,1
	DIVTR"显示位置导航：","是否显示[你的位置]导航条<br />(包括版块下拉菜单)","<br>"&GetRadio("info27","否",Info(27),0)&GetRadio("info27","是",Info(27),1),60,2
'30-39上传设置
	.Write"</div><a name='inc3'></a><div class='mian'><div class='top'>上传设置<a href='#inc'>▲</a></div>"
	DIVTR"文件上传：","是否允许开启上传",GetRadio("info30","禁止",Info(30),0)&GetRadio("info30","开启",Info(30),1),40,1
	DIVTR"防盗链：","是否开启上传文件防盗链",GetRadio("info31","否",Info(31),0)&GetRadio("info31","是",Info(31),1),40,2
	DIVTR"文件下载计数：","是否开启上传文件的下载计数器",GetRadio("info32","否",Info(32),0)&GetRadio("info32","是",Info(32),1),40,1
	DIVTR"采用数据流下载或显示：","是否采用数据流组件（隐藏式）",GetRadio("info38","否",Info(38),0)&GetRadio("info38","是",Info(38),1),40,1
	DIVTR"头像上传大小：","限制头像最大上传的大小","<input type='text' class='text' name='info33' size='5' value='"&Info(33)&"' /> KB",40,2
	DIVTR"上传文件类型：","允许上传的可以下载的类型，每个字符用“|”隔开","<input type='text' class='text' name='info34' size=60 style='WIDTH: 99%';' value='"&Info(34)&"' />",55,1
	DIVTR"上传图片类型：","允许上传的可以显示的图片类型，每个字符用“|”隔开","<input type='text' class='text' name='info35' size=60 style='WIDTH: 99%';' value='"&Info(35)&"' />",55,2
	DIVTR"文件上传目录：","如果更改,需同时通过FTP新建目录和移动原来的文件","<input type='text' class='text' name='info36' size='20'  value='"&Info(36)&"' />",55,1	
	DIVTR"头像上传目录：","如果更改,需同时通过FTP新建目录和移动原来的文件","<input type='text' class='text' name='info37' size='20' value='"&Info(37)&"' />",55,2
	DIVTR"上传信息边框：","是否在帖子上显示上传的信息边框？",GetRadio("info39","否",Info(39),0)&GetRadio("info39","是",Info(39),1),55,2

'40-59用户选项
	.Write"</div><a name='inc4'></a><div class='mian'><div class='top'>用户选项<a href='#inc'>▲</a></div>"
	DIVTR"用户注册：","是否允许用户注册？",GetRadio("info40","否",Info(40),0)&GetRadio("info40","是",Info(40),1),40,1
	DIVTR"注册审核：","用户注册的帐号是否要通过审核才能使用.",GetRadio("info41","否",Info(41),0)&GetRadio("info41","是",Info(41),1),40,2
	DIVTR"注册邮箱限制：","是否设定一个邮箱只能注册一个帐号.",GetRadio("info42","否",Info(42),0)&GetRadio("info42","是",Info(42),1),40,1
	DIVTR"注册欢迎留言：","用户注册完,是否自动发送站内欢迎留言.",GetRadio("info43","否",Info(43),0)&GetRadio("info43","是",Info(43),1),40,2
	DIVTR"欢迎留言内容：","用户注册完，自动发送的留言。","<textarea rows='3' name='info46' cols='70'>"&Info(46)&"</textarea>",55,1
	DIVTR"允计个人签名：","是否允许帖子显示用户签名.",GetRadio("info44","否",Info(44),0)&GetRadio("info44","是",Info(44),1),40,1
	DIVTR"新留言提示：","当用户有新留言时的提示方式.",GetRadio("info45","声音/图标",Info(45),0)&GetRadio("info45","窗口弹出",Info(45),1),40,2
	DIVTR"在线用户分页个数：","设置在线用户列表的每页显示人数","<input type='text' class='text' name='info47' size='2' value='"&Info(47)&"' /> ",40,2
	DIVTR"版主继承：","设定上级版主可以管理下级子论坛.",GetRadio("info48","否",Info(48),0)&GetRadio("info48","是",Info(48),1),40,1
	DIVTR"限制每人每日的评帖次数：","可以有效减少滥用权力（站长不受限制）","<input type='text' class='text' name='info49' size='5' value='"&Info(49)&"' /> 次",40,2
'	DIVTR"版主评帖：","版主是否可以进行评帖和奖惩操作？",GetRadio("info50","否",Info(50),0)&GetRadio("info50","是",Info(50),1),40,1
	DIVTR"删除帖子操作选项：","当删除别人帖子时，是否显示选项？",GetRadio("info51","否",Info(51),0)&GetRadio("info51","是",Info(51),1),40,2
	DIVTR"禁止注册的用户名：","用于过滤用户名和头衔称号,用“|”隔开","<input type='text' class='text' name='info52' style='WIDTH: 99%';'size='60' value='"&Info(52)&"' />",40,1	
	DIVTR"论坛头像个数：","设定论坛自带头像的数目","<input type='text' class='text' name='info53' size='5' value='"&Info(53)&"' /> 个",40,2
	DIVTR"头像默认宽度：","头像的默认寸尺宽度","<input type='text' class='text' name='info54' size='5' value='"&Info(54)&"' /> px",40,1
	DIVTR"头像默认高度：","头像的默认寸尺高度","<input type='text' class='text' name='info55' size='5' value='"&Info(55)&"' /> px",40,2
	DIVTR"头像最大尺寸：","限制头像最大高度和宽度","<input type='text' class='text' name='info56' size='5' value='"&Info(56)&"' /> px",40,1
	DIVTR"外部头像图片：","是否开启用户头像可以外部连接图片？",GetRadio("info57","禁止",Info(57),0)&GetRadio("info57","开启",Info(57),1),40,2
'60-79帖子设置
	.Write"</div><a name='inc5'></a><div class='mian'><div class='top'>帖子设置<a href='#inc'>▲</a></div>"
	DIVTR"发帖模式：","设定发帖编辑器",GetRadio("info60","HTML(全功能模式)",Info(60),0)&GetRadio("info60","UBB(数据节省模式)",Info(60),1),40,1
	DIVTR"主题列表条数：","主题列表(board.asp)每页的显示条数","<input type='text' class='text' name='info61' size='5' value='"&Info(61)&"' />",40,2
	DIVTR"帖子回复条数：","帖子显示(Topic.asp)每页的显示条数","<input type='text' class='text' name='info80' size='5' value='"&Info(80)&"' />",40,1
	DIVTR"帖子打开窗口：","主题列表的打开方式",GetRadio("info69","原窗口",Info(69),0)&GetRadio("info69","新窗口",Info(69),1),40,2
	DIVTR"热帖标准：","成为热门主题的回复帖数","<input type='text' class='text' name='info62' size='5' value='"&Info(62)&"' />",40,2
	DIVTR"投票条数：","用户发投票主题的最大数目","<input type='text' class='text' name='info63' size='5' value='"&Info(63)&"' />",40,1
	DIVTR"游客查看精华帖 ：","是否允许游客浏览精华主题",GetRadio("info64","否",Info(64),0)&GetRadio("info64","是",Info(64),1),40,2
	DIVTR"开启贴图：","是否在帖子识别UBB图片标签",GetRadio("info65","否",Info(65),0)&GetRadio("info65","是",Info(65),1),40,1
	DIVTR"开启识别连接：","是否开启自动识别帖子上的网址连接？",GetRadio("info82","否",Info(82),0)&GetRadio("info82","是",Info(82),1),40,2
	DIVTR"开启Flash：","是否在帖子识别UBB动画标签",GetRadio("info66","否",Info(66),0)&GetRadio("info66","是",Info(66),1),40,1
	DIVTR"开启播放器：","是否识别UBB音乐视频播放器MP/RM",GetRadio("info67","否",Info(67),0)&GetRadio("info67","是",Info(67),1),40,2
	DIVTR"开启插入代码：","是否开启识别代码转换标签",GetRadio("info68","否",Info(68),0)&GetRadio("info68","是",Info(68),1),40,1
	DIVTR"特殊帖_回复可见：","是否开启发表只有回复主题可见的特殊帖子",GetRadio("info70","否",Info(70),0)&GetRadio("info70","是",Info(70),1),40,1
	DIVTR"特殊帖_金钱可见：","是否开启发表达到指定金钱数量可见的特殊帖子",GetRadio("info71","否",Info(71),0)&GetRadio("info71","是",Info(71),1),55,2
	DIVTR"特殊帖_积分可见：","是否开启发表达到指定积分可见的特殊帖子",GetRadio("info72","否",Info(72),0)&GetRadio("info72","是",Info(72),1),55,1
	DIVTR"特殊帖_日期可见：","是否开启发表在指定日期后可见的特殊帖子",GetRadio("info73","否",Info(73),0)&GetRadio("info73","是",Info(73),1),40,2
	DIVTR"特殊帖_性别可见：","是否开启发表指定用户性别可见的特殊帖子",GetRadio("info74","否",Info(74),0)&GetRadio("info74","是",Info(74),1),40,1
	DIVTR"特殊帖_登陆可见：","是否开启发表只有在登陆后可见的特殊帖子",GetRadio("info75","否",Info(75),0)&GetRadio("info75","是",Info(75),1),40,2
	DIVTR"特殊帖_指定读者：","是否开启发表只有指定会员可见的特殊帖子",GetRadio("info76","否",Info(76),0)&GetRadio("info76","是",Info(76),1),40,1
	DIVTR"特殊帖_付费观看：","是否开启发表可以挣取金钱的特殊帖子",GetRadio("info77","否",Info(77),0)&GetRadio("info77","是",Info(77),1),40,2
	DIVTR"发帖后提示：","是否开启用户发帖后显示奖励的信息？",GetRadio("info78","否",Info(78),0)&GetRadio("info78","是",Info(78),1),40,1
	DIVTR"帖子过滤脏字：","帖子字符过滤后用被“*”代替<br>每个字符请用“|”隔开","<br /><input type='text' class='text' name='info79' style='WIDTH: 99%';' value='"&Info(79)&"' />",55,2	
	DIVTR"最后回复显示：","当用户回复帖子在各版块的显示？",GetRadio("info81","主题的标题",Info(81),0)&GetRadio("info81","回复的内容",Info(81),1),40,1
'90-论坛奖励
	.Write"</div><a name='inc6'></a><div class='mian'><div class='top'>论坛资源分配<a href='#inc'>▲</a></div>"
	DIVTR"资源别名：","论坛的三个奖励参数值，可以根据论坛的需要改成其它名称。<br>例如：可以把“积分”改为“威望”等","<input type='text' class='text' name='info120' size='20' value='"&Info(120)&"' /> 默认名称：金钱<br /><input type='text' class='text' name='info121' size='20' value='"&Info(121)&"' /> 默认名称：积分<br /><input type='text' class='text' name='info122' size='20' value='"&Info(122)&"' /> 默认名称：游戏币",58,1
	DIVTR"总置顶奖励：","当被设为总置顶主题对作者的奖励，解除则减少相应资源",Info(120)&"：<input type='text' class='text' name='info90' size='5' value='"&Info(90)&"' /> "&Info(121)&"：<input type='text' class='text' name='info91' size='5' value='"&Info(91)&"' /> "&Info(122)&"：<input type='text' class='text' name='info92' size='5' value='"&Info(92)&"' />",58,2
	DIVTR"区置顶奖励：","当被设为区置顶主题对作者的奖励，解除则减少相应资源",Info(120)&"：<input type='text' class='text' name='info93' size='5' value='"&Info(93)&"' /> "&Info(121)&"：<input type='text' class='text' name='info94' size='5' value='"&Info(94)&"' /> "&Info(122)&"：<input type='text' class='text' name='info95' size='5' value='"&Info(95)&"' />",58,1
	DIVTR"置顶奖励：","当被设为置顶主题对作者的奖励，解除则减少相应资源",Info(120)&"：<input type='text' class='text' name='info96' size='5' value='"&Info(96)&"' /> "&Info(121)&"：<input type='text' class='text' name='info97' size='5' value='"&Info(97)&"' /> "&Info(122)&"：<input type='text' class='text' name='info98' size='5' value='"&Info(98)&"' />",58,2
	DIVTR"精华奖励：","当被设为精华主题对作者的奖励，解除则减少相应资源",Info(120)&"：<input type='text' class='text' name='info99' size='5' value='"&Info(99)&"' /> "&Info(121)&"：<input type='text' class='text' name='info100' size='5' value='"&Info(100)&"' /> "&Info(122)&"：<input type='text' class='text' name='info101' size='5' value='"&Info(101)&"' />",58,1
	DIVTR"发表主题奖励：","用户发表主题帖的奖励",Info(120)&"：<input type='text' class='text' name='info102' size='5' value='"&Info(102)&"' /> "&Info(121)&"：<input type='text' class='text' name='info103' size='5' value='"&Info(103)&"' /> "&Info(122)&"：<input type='text' class='text' name='info104' size='5' value='"&Info(104)&"' />",40,2
	DIVTR"发表回复奖励：","用户发表回复帖的奖励",Info(120)&"：<input type='text' class='text' name='info105' size='5' value='"&Info(105)&"' /> "&Info(121)&"：<input type='text' class='text' name='info106' size='5' value='"&Info(106)&"' /> "&Info(122)&"：<input type='text' class='text' name='info107' size='5' value='"&Info(107)&"' />",40,1
	DIVTR"删除惩罚：","当帖子被删除时对作者的默认惩罚",Info(120)&"：<input type='text' class='text' name='info108' size='5' value='"&Info(108)&"' /> "&Info(121)&"：<input type='text' class='text' name='info109' size='5' value='"&Info(109)&"' /> "&Info(122)&"：<input type='text' class='text' name='info110' size='5' value='"&Info(110)&"' />",58,2
	DIVTR"回复主题奖励：","每次回复同时给主题作者的奖励",Info(120)&"：<input type='text' class='text' name='info111' size='5' value='"&Info(111)&"' /> ",58,1
	DIVTR"发帖字符少不奖励：","设定小于此字符数将不会给于奖励","发帖字符数：<input type='text' class='text' name='info112' size='5' value='"&Info(112)&"' /> ",58,2
	DIVTR"留言收费：","当用户发送留言时扣取费用","<input type='text' class='text' name='info123' size='5' value='"&Info(123)&"' /> 金钱",40,1
	DIVTR"评帖最大限：","设定在评帖时进行奖和罚操作的最大限度",Info(120)&"：<input type='text' class='text' name='info113' size='5' value='"&Info(113)&"' /> "&Info(121)&"：<input type='text' class='text' name='info114' size='5' value='"&Info(114)&"' /> "&Info(122)&"：<input type='text' class='text' name='info115' size='5' value='"&Info(115)&"' />",58,1
	DIVTR""&Info(121)&"汇率：",""&Info(121)&"的汇率","1000个"&Info(120)&" = <input type='text' class='text' name='info116' size='5' value='"&Info(116)&"' /> 个"&Info(121)&"",40,2
	DIVTR""&Info(122)&"汇率：",""&Info(122)&"的汇率","1000个"&Info(120)&" = <input type='text' class='text' name='info117' size='5' value='"&Info(117)&"' /> 个"&Info(122)&"",40,1
	DIVTR"版主月薪：","版主每月的工资","<input type='text' class='text' name='info118' size='5' value='"&Info(118)&"' /> "&Info(120)&"",40,2
	DIVTR"银行利率：","用户储存在银行的"&Info(120)&"每日利率","<input type='text' class='text' name='info119' size='5' value='"&Info(119)&"' /> %",40,1

	.Write"<div class='bottom'><input type='submit' class='button' value='确定修改'><input type='reset' class='button' value='重 置'><a href='#inc'>▲</a></div></div></form>"
	End with
End Sub


Sub AddMenu
	Dim ParenID,S,Rs
	ParenID=Request("ParenID")
	Response.Write GoForm("SaveMenu")
	Response.Write"<div class='mian'><div class='top'>添加论坛导航菜单</div>"
	DIVTR"名称：","","<input type='text' class='text' name='MenuName' size='20' /> *",25,1
	DIVTR"连接文件：","","<input type='text' class='text' name='MenuUrl' size='35' />(请填写相对路径,留空则不连接。)",25,2
	DIVTR"所属菜单：","",MenuSelect(ParenID),25,1
	DIVTR"显示可见：","","<select name='Show'><option value='0' selected>全部可见</option><option value='1'>只有会员可见</option><option value='2'>只有游客可见</option><option value='3'>不可见(隐藏)</option></select>",25,2
	DIVTR"打开方式：","","<select name='Target'><option value='0' selected>原窗口</option><option value='1'>新窗口</option></select>",25,1
	Response.Write"<div class='bottom'><input type='submit' value=' 提 交 '>&nbsp;&nbsp;<input type='reset' value=' 重 置 '></div></div></form>"
End Sub

Sub EditMenu
	Dim ID,Rs,S
	ID=request.querystring("ID")
	Set Rs=BBS.Execute("Select name,Url,Show,Flag,ParenID,Target From [Menu] where ID="&ID&"")
	If Rs.Eof Then Goback"","记录不存在"
	Response.Write GoForm("SaveMenu")
	If Rs(3)>0 Then S="系统<input name='Flag' type='hidden' value='"&Rs(3)&"' />" Else S="普通"
	Response.Write"<div class='mian'><div class='top'>修改论坛"&S&"菜单</div>"
	DIVTR"名称：","","<input name='ID' type='hidden' value='"&ID&"' /><input name='MenuName' type='text' class='text' value='"&Rs(0)&"' size='20'> *",25,1
	If Rs(3)>0 Then
		S=Rs(1)
	Else
		S="<input name='MenuUrl' type='text' class='text' value='"&Rs(1)&"' size='38' />(请填写相对路径,留空则不连接。)"
	End If
	DIVTR"连接文件：","",S,25,1
	If Rs(3)<>8 Then
	DIVTR"所属菜单：","",MenuSelect(Rs(4)),25,1
	DIVTR"打开窗口：","",GetRadio("Target","原窗口",Rs(5),0)&GetRadio("Target","新窗口",Rs(5),1),25,1
	End If
	DIVTR"显示可见：","",GetRadio("Show","全部可见",Rs(2),0)&GetRadio("Show","只有会员可见",Rs(2),1)&GetRadio("Show","只有游客可见",Rs(2),2)&GetRadio("Show","不可见(隐藏)",Rs(2),3),25,1
	Response.Write"<div class='bottom'><input type='submit' value=' 提 交 '>&nbsp; <input type='reset' value=' 重 置 '></div></div></form>"
	Rs.Close
End Sub

Sub Menu
	Dim Showmood,Sql,Rs1,Subs,I,S
	With Response
	Showmood=Request("Showmood")
	.Write GoForm("MenuOrder")&"<div class='mian'><div class='top'>论坛菜单</div><div class='divth' style=';padding:5px'><div style='FLOAT: right;'>查看方式：<a href='?Action=Menu&Showmood=2'>游客菜单</a> <a href='?Action=Menu&Showmood=1'>会员菜单</a> <a href='?Action=Menu'>显示全部</a></div><div>【<a href='?Action=AddMenu&ParenID=0'>"&IconA&"添加菜单</a>】 【<a href='Admin_Confirm.asp?action=setjsmenu'>生成菜单</a>】</div></div>"
    .Write"<table class='Stable' border='0' cellpadding='3' cellspacing='0'><tr><th width='50px'>排序</th><th width='20%'>名称</th><th width='30%'>连接文件</th><th width='55px'>显示</th><th width='30px'>属性</th><th>操作</th></tr>"
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
		.Write"<font color=red>风格</font>"
	ElseIf Rs(5)>0 Then
		.Write"系统"
	Else
		.Write"普通"
	End If
	.Write"</td><td><a href='?Action=EditMenu&ID="&Rs(0)&"'>"&IconE&"编辑</a> "
	Subs=BBS.Execute("Select Count(*) From [Menu] where parenID="&Rs(0))(0)
	If Rs(5)=0 then
	.Write"<a href=""javascript:"
	If Subs>0 Then
		.Write"alert('该菜单有下拉项目，不能删除，请先移除属下的下拉菜单项目。')"">"
	Else
        .Write"checkclick('删除后将不能恢复！您确定要删除吗？','Admin_Confirm.asp?Action=DelMenu&ID="&Rs(0)&"')"" >"
	End If 
	.Write IconD&"删除</a> "
	End IF
	If Rs(5)<>8 Then .Write "<a href='?Action=AddMenu&ParenID="&Rs(0)&"'>"&IconA&"添加下拉项</a>"
	.Write"</td></tr>"
	'风格菜单-只读
	If Rs(5)=8 Then
		Set Rs1=BBS.Execute("Select SkinID,SkinName,IsDefault,Ismode,Pass,remark From [Skins] Order By SkinID Asc")
		Do while not Rs1.eof
			.Write"<tr><td>├</td>"
			.Write"<td>"&Rs1(1)&"</td><td>&nbsp;"
			.Write"</td><td>"&MenuShow(Rs(3))&"</td><td>"
			.Write"风格</td><td><a href=""javascript:checkclick('此菜单为系统风格，需在风格管理中编辑。\n现在要转到风格管理中编辑吗？','Admin_Template.asp')"">"&IconE&"编辑</a></td></tr>"
		Rs1.movenext
		Loop
		Rs1.Close
	End If
	'下拉菜单
	If Subs>0 Then
		If ShowMood="" Then
			S="parenID="&Rs(0)&" order by orders"
		Else
			S="parenID="&Rs(0)&" and (Show="&showmood&" or Show=0) order by orders"
		End If
		Set Rs1=BBS.Execute(Sql&S)
		Do while Not Rs1.eof
			.Write"<tr><td>├<input name='Orders' type='text' class='text' value='"&Rs1(4)&"' size='2'><input name='ID' type='hidden' value='"&Rs1(0)&"'></td>"
			.Write"<td>"&Rs1(1)&"</td><td>"
			If Rs1(2)<>"" Then .Write"<a href='"&Rs1(2)&"' target=_blank>"&Rs1(2)&"</a>" Else .Write "&nbsp;"
			.Write"</td><td>"&MenuShow(Rs1(3))&"</td><td>"
			If Rs1(5)>0 Then
				.Write"系统"
			Else
				.Write"普通"
			End If
			.Write"</td><td><a href='?Action=EditMenu&ID="&Rs1(0)&"'>"&IconE&"编辑</a> "
			If Rs1(5)=0 then
			.Write"<a href=""javascript:checkclick('删除后将不能恢复！您确定要删除吗？','Admin_Confirm.asp?Action=DelMenu&ID="&Rs1(0)&"')"">"&IconD&"删除</a></td></tr>"
			End If
		Rs1.movenext
		Loop
		Rs1.Close
	End If
	Rs.Movenext
	Loop
	Rs.Close
	Set Rs1=nothing
	.Write"</table><div class='bottom'><input type='submit' class='button' value='更新排序' /></div></div></form>"
	End With
End Sub

Function MenuShow(Show)
	Select case Show
	case "1"
	MenuShow="只有会员"
	Case "2"
	MenuShow="只有游客"
	Case "3"
	MenuShow="<font color=blue>不显示</font>"
	Case else
	MenuShow="全显示"
	End Select
End Function
Function MenuSelect(parenID)
	Dim mRs,Temp
	Temp="<select name='ParenID'><option value='0'>..做为菜单导航</option>"
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
	Response.Write GoForm("UpdateConfigData")&"<div class='mian'><div class='top'>论坛系统数据设置</div><div class='divth'>说明：以下信息一般不建议用户修改，带*号的在整理论坛时将会被自动修正</div>"
	DIVTR"论坛会员总数 ：","论坛注册用户总数","<input type='text' name='usernum' size='20' class='text' value='"&.InfoUpdate(5)&"'> *" ,40,1
	DIVTR"论坛帖子总数 ：","论坛所有帖子总数","<input type='text' name='allessaynum' size='20' class='text' value='"&.InfoUpdate(0)&"'> *" ,40,2
	DIVTR"论坛主题总数 ：","论坛主题帖子总数","<input type='text' name='topicnum' size='20' class='text' value='"&.InfoUpdate(1)&"'> *" ,40,1
	DIVTR"论坛最高日发贴：","记录历史最高的日发贴","<input type='text' name='maxessaynum' size='20' class='text' value='"&.InfoUpdate(4)&"'> " ,40,2
	DIVTR"最高在线人数：","历史最高同时在线纪录人数","<input type='text' name='maxonlinenum' size='20' class='text' value='"&.InfoUpdate(7)&"'> " ,40,1
	DIVTR"最高在线人数发生时间：","历史最高同时在线纪录人数的那个时间","<input type='text' name='maxonlinetime' size='20' class='text' value='"&.InfoUpdate(8)&"'> (格式：YYYY-M-D H:M:S)" ,40,2
	DIVTR"论坛执行次数：","页面下部的计数器","<input type='text' name='hits' size='20' class='text' value='"&.InfoUpdate(9)+Temp&"'>" ,40,1
	Response.Write"<div class='bottom'><input type='submit' value='确定修改' class='button'><input class='button' type='reset' value='重 置'></div></div></form>"
	End With
End Sub

Sub A_E_LockIP
	Dim ID,StartIP,EndIp,Readme,Title
	ID=request.querystring("ID")
	StartIP=request.querystring("IP")
	Readme=request.querystring("Readme")
	Title="IP封锁"
	If ID<>0 Then
		Set Rs=BBS.execute("Select StartIp,EndIp,Readme,ID From[LockIp] where ID="&ID&"")
		IF Rs.eof Then
			GoBack"","记录不存在"
			Exit Sub
		Else
			Title="修改封锁IP"
			StartIP=BBS.Fun.IpDeCode(Rs(0))
			EndIp=BBS.Fun.IpDeCode(Rs(1))
			Readme=Rs(2)
		End If
	End If
	Response.Write GoForm("LockIp")&"<div class='mian'><div class='top'>"&Title&"</div><input name='ID' type='hidden' value='"&ID&"' />"
	DIVTR"起始IP：","此项必需填写","<input name='StartIp' type='text' class='text' value='"&StartIp&"' />",35,1
	DIVTR"结束IP：","封锁单个IP时不必填写","<input name='EndIp' type='text' class='text' value='"&EndIp&"' />",35,1
	DIVTR"封禁说明：","最大255个字符","<input name='Readme' type='text' class='text' style='width:90%' value='"&Readme&"' />",35,1
	Response.Write"<div class='bottom'><input type='submit' class='button' value='确 定' /><input type='reset' class='button' value='重 置' /></div></div></form>"
End Sub

Sub LockIp
	Dim S
	A_E_LockIP()
	Response.Write"<div class='mian'><div class='top'>已经被封的IP记录</div>"
	S="<table border='0' class='Stable' cellpadding='3' cellspacing='0'><tr><th width='35%'>网段</th><th width='40%'>说明</th><th>操作</th></tr>"
	Set Rs=BBS.Execute("Select StartIp,EndIp,Readme,Lock,ID From[LockIp] where Lock=1")
	If Rs.eof Then
		Response.Write"<div class='divtr1'>没有封锁记录</div>"
	Else
		Response.Write S
		Do while not Rs.eof
			Response.Write"<tr><td>"&BBS.Fun.IpDeCode(Rs("StartIp"))&" ～ "&BBS.Fun.IpDeCode(Rs("EndIp"))&"</td><td>&nbsp;"&Rs("Readme")&"</td><td align='center'><a href=?Action=EditIp&Id="&rs("ID")&">"&IconE&"修改<a> <a href=Admin_Confirm.asp?Action=IsLockIp&ID="&rs("ID")&"><img src='Images/icon/lock.gif' align='absmiddle' border='0' /> 解除</a></td></tr>"
		Rs.MoveNext
		Loop
		Response.Write"</table>"
	End If
	Rs.Close
	Response.Write"</div>"
	Response.Write"<div class='mian'><div class='top'>未被封的IP记录</div>"
	Set Rs=BBS.Execute("Select StartIp,EndIp,Readme,Lock,ID From[LockIp] where Lock=0")
	If Rs.eof Then
		Response.Write"<div class='divtr1'>没有记录</div>"
	Else
	Response.Write S
	Do while not Rs.eof
		Response.Write"<tr><td>"&BBS.Fun.IpDeCode(Rs("StartIp"))&" ～ "&BBS.Fun.IpDeCode(Rs("EndIp"))&"</td><td>&nbsp;"&Rs("Readme")&"</td><td align='center'><a href=Admin_Confirm.asp?Action=IsLockIp&ID="&Rs("ID")&"><img src='Images/icon/lock.gif' border=0 align='absmiddle' /> 封锁</a> <a href=#this onclick=""checkclick('删除后将不能恢复！您确定要删除吗？','Admin_Confirm.asp?Action=DelLockIP&ID="&Rs("ID")&"')"">"&IconD&"删除</td></tr>"
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
	.Write"<div class='mian'><div class='top'>数据表管理</div><div class='divtr2' style='padding:5px'><b>说明：</b><br />默认数据表：默认选中的为当前论坛所使用来保存帖子数据的表。<br>删除数据表：删除会同时全部删除该数据表的所有帖子，请注意！！！<br>增加数据表：当帖子数量非常多时，建议(Access版本用户每个表超过5万左右，SQL版本用户每个表超过25万左右)添加一个数据表。<br>合并数据表：合并后，“指定数据表”会被删除，所有的帖子会移动到“目标数据表”中，默认表不能做为“指定数据表”。</div></div>"
	.Write GoForm("AuteSqlTable")&"<div class='mian'><div class='top'>设置默认数据表</div><table class='Stable' border='0' cellpadding='3' cellspacing='0'>"&_
	"<tr><th width='100px'>数据表</th><th width='100px'>帖数</th><th width='10%'>默认</th><th>操作</th></tr>"
	AllTable=Split(BBS.BBStable(0),",")
	For i=0 To uBound(AllTable)
		.Write"<tr><td>Bbs"&AllTable(i)&"</td><td>"&BBS.execute("Select Count(*) From[bbs"&AllTable(i)&"]")(0)&"</td><td><input name='Aute' type='radio' value='"&AllTable(i)&"'"
		If BBS.BBStable(1)=AllTable(i) Then
			.Write" checked /></td><td><a onclick=alert('该数据表为默认数据表，不能删除默认的数据表！') href='#this'>"
		Else
			.Write" /></td><td><a onclick=""checkclick('注意！删除将包括数据表的所有帖子！\n\n删除后将不能恢复！您确定要删除吗？','Admin_Confirm.asp?Action=DelSqlTable&ID="&AllTable(i)&"')""  href='#'>"
		End If
		.Write IconD&"删除</a></td></tr>"
	Next
	.Write"</table><div class='bottom'><input type='submit' value='提 交' class='button' /><input type='reset' value='重 置'  class='button' /></div></div></form>"
	.Write GoForm("AddSqlTable")&"<div class='mian'><div class='top'>增加数据表</div><div class='divtr1' style='padding:5px'>新数据表名称：bbs<input type='text' name='TableName' class='text' size='2' value='"&uBound(AllTable)+2&"' ONKEYPRESS='event.returnValue=(event.keyCode >= 48) && (event.keyCode <= 57);' /> (只填写数字，不能和现有的数据表相同。)</div>"
	.Write"<div class='bottom'><input type='submit' value='提 交' class='button' /></div></div></form>"
	.Write GoForm("SqlTableUnite")&"<div class='mian'><div class='top'>合并数据表</div><div class='divtr1' style='padding:5px'>将数据表：<select name='SqlTableID1'><option value='0'>指定数据表</option>"&GetSqlTableList&"</select> 所有的帖子合并到：数据表 <select name='SqlTableID2'><option value='0'>目标数据表</option>"&GetSqlTableList&"</select>中！ </div>"
	.Write"<div class='bottom'><input type='button' value='提 交' class='button' onclick=""if(confirm('注意！操作后将不能恢复！您确定要合并数据表吗？'))form.submit()"" /></div></div></form>"
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
			Goback"","空间不支持FOS文件读写！。"
			err.Clear
			Exit Sub
		End If
	Set fso=nothing
	Response.Write"<div class='mian'><div class='top'>系统空间占用情况</div>"
	DIVTR"论坛数据占用空间：","","<img src='Images/icon/hr6.gif' style='margin-top:8px' width='"&drawbar("data")&"' height='10' alt='' /> "&GetSpaceinfo("data"),25,1
	DIVTR"备份数据占用空间：","","<img src='Images/icon/hr6.gif' style='margin-top:8px' width='"&drawbar("data_backup")&"' height='10' alt='' /> "&GetSpaceinfo("data_backup"),25,2
	DIVTR"程序文件占用空间：","","<img src='Images/icon/hr3.gif' style='margin-top:8px' width='"&drawbar("i@BBS@")&"' height='10' alt='' /> "&GetSpaceinfo("i@BBS@"),25,1
	DIVTR"Inc 目录占用空间：","","<img src='Images/icon/hr3.gif' style='margin-top:8px' width='"&drawbar("inc")&"' height='10' alt='' /> "&GetSpaceinfo("inc"),25,2
	DIVTR"图片目录占用空间：","","<img src='Images/icon/hr5.gif' style='margin-top:8px' width='"&drawbar("pic")&"' height='10' alt='' /> "&GetSpaceinfo("pic"),25,1
	DIVTR"皮肤目录占用空间：","","<img src='Images/icon/hr4.gif' style='margin-top:8px' width='"&drawbar("skins")&"' height='10' alt='' /> "&GetSpaceinfo("skins"),25,2
	DIVTR"上传头像占用空间：","","<img src='Images/icon/hr2.gif' style='margin-top:8px' width='"&drawbar("UploadFile/Head")&"' height='10' alt='' /> "&GetSpaceinfo("UploadFile/Head"),25,1
	DIVTR"上传文件占用空间：","","<img src='Images/icon/hr2.gif' style='margin-top:8px' width='"&drawbar("UploadFile/TopicFile")&"' height=10 alt='' /> "&GetSpaceinfo("UploadFile/TopicFile"),25,2
	Response.Write"<div class='bottom' style='padding:2px'>论坛占用空间总计：<img src='Images/icon/hr1.gif' width='400' height='10' alt='' style='margin-top:8px' /> "&GetSpaceinfo("i@BBS")&"</div></div>"
End Sub
'2005-12-25重写 by suibing
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
	Response.Write"<div class='mian'><div class='top'>论坛整理修复</div>"&_
	"<div class='divth' style='text-align :left;padding:5px'>注意事项：论坛整理中的各项运行都可能非常消耗服务器资源，时间也可能很长，请耐心等候。<br>所以请你选择论坛访问人数较少的时候进行整理， 或者在整理过程中可以先暂时【<a href=?Action=BbsInfo>关闭论坛</a>】</div>"&_
	"<div class='divtr1' style='padding:5px'><b>论坛系统整理</b><br>重新计算总主题数、总帖数、今日帖数、用户数、新注册用户等，建议每隔一段时间运行一次。<br /><input value='开始整理' type='button' class='button' onClick=window.location.href='Admin_Confirm.asp?action=UpdateBbsdate' /></div>"&_
	"<div class='divtr2' style='padding:5px'><b>论坛版面整理</b><br />重新计算论坛各版面总帖数、主题数、今日帖数、各版版主、最后回复等，建议每隔一段时间运行一次。清理的过程中请不要刷新和关闭！<br /><input value='开始整理' type='button' class='button' onClick=window.location.href='Admin_Board.asp?Action=BoardUpdate' /></div>"&_
	"<div class='divtr1' style='padding:5px'><b>论坛垃圾清理</b><br />清理无效版主、无效帖子、无效主题、无效帖子、无效投票、无效留言、无效用户帖等，整理过程可能将消耗大量资源，建议在本地上进行，清理的过程中请不要刷新和关闭！<br /><input value='开始清理' type='button' class='button' onClick=window.location.href='Admin_Confirm.asp?Action=DelWuiong' /></div>"&_
	"<div class='divtr2' style='padding:5px'><b>修复主题帖数</b><br />重新整理计算每个主题帖的回复帖数、最后回复信息等，如果论坛帖子非常多，整理过程可能将消耗大量资源。<br /><input value='开始整理' type='button' class='button' onClick=window.location.href='Admin_Confirm.asp?Action=UpdateTopic' /></div>"&_
	"<div class='divtr1' style='padding:5px'><b>修复用户信息</b><br />重新整理计算每个用户的等级、总帖数、精华帖数等，如果注册会员非常多，整理过程可能将消耗大量资源。<br /><input value='开始整理' type='button' class='button' onClick=window.location.href='Admin_Confirm.asp?action=UpdateAllUser' /></div></div>"
End Sub

Sub A_E_Link
	Dim Title,ID,Orders,Ispic,Pic,BbsName,Admin,Url,Readme,Pass
	pass=1
	Ispic=0
	Title="添加"
	ID=Request("ID")
	If ID<>"" Then
		Set Rs=BBS.Execute("Select ID,Orders,IsPic,Pic,BbsName,Admin,Url,Readme,pass From [Link] where ID="&ID&"")
		IF Rs.eof Then
			GoBack"","这条论坛联盟不存在！"
			Exit Sub
		Else
			Title="修改"
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
	Response.Write"<div class='mian'><div class=top>"&Title&"修改论坛联盟</div>"
	DIVTR "论坛名称：","","<input name='ID' value='"&ID&"' type='hidden'><input type='text' class='text' name='bbsname' size='15' value='"&BbsName&"' />",25,1
	DIVTR "论坛地址：","","<input type='text' name='url' size='28' class='text' value='"&Url&"' />",25,2
	DIVTR "论坛站长：","","<input type='text' name='admin' size='20' class='text' value='"&Admin&"' />(可以留空)",25,1
	DIVTR "论坛图片：","","<input type='text' name='pic' size='38' class='text' value='"&Pic&"'>(即使不做图片连接-也必须填写{可随便乱写})",25,2
	DIVTR "论坛说明：","","<input type='text' name='Readme' size='38' class='text' value='"&Readme&"'>(可以留空)",25,1
	DIVTR "图片显示：","",GetRadio("ispic","否",ispic,0)&GetRadio("ispic","是",ispic,1),25,2
	DIVTR "通过审核：","",GetRadio("pass","×",pass,0)&GetRadio("pass","<font color=red>√ </font>",pass,1),25,1
	Response.Write"<div class='bottom'><input type='submit' class='button' value='提 交'><input type='reset' class='button' value='重 置'></div></div></form>"
End Sub

Sub Grade()
	Dim Arr_Rs,i,S
	Response.Write GoForm("AllUpdateGrade")&"<input name='Grouping' type='hidden' value='0' /><div class='mian'><div class='top'>用户等级设置</div><table border='0' class='Stable' cellpadding='3' cellspacing='0'><tr><th width='15%'>等级名称</th><th width='10%'>所需帖数</th><th width='30%'>等级图片</th><th width='19%'>标志图片</th><th width='20%'>管理操作</th></tr>"
	Set Rs=BBS.execute("Select Grouping,ID,GradeName,EssayNum,PIC,Spic,Flag FROM [Grade] where Grouping=0 order by EssayNum")
	do while Not Rs.Eof
	Response.Write"<tr><td align='center'><input name='ID' type='hidden' value='"&Rs(1)&"' />"&_
	"<input class='text' name='GradeName' type='text' size='15' value='"&Rs(2)&"' /></td>"&_
	"<td align='center'><input class='text' name='EssayNum' type='text' size='4' value='"&Rs(3)&"' /></td><td><input class='text' name='Pic' type='text' size='15' value='"&Rs(4)&"' /><img src='Pic/Grade/"&Rs(4)&"' /></td>"&_
	"<td><input class='text' name='Spic' type='text' size='15' value='"&Rs(5)&"' />"
	If len(Rs(5))>3 Then Response.Write"<img src='Pic/Grade/"&Rs(5)&"' /></td>"
	Response.Write "<td>"&IconE&"<a href='?action=A_E_Grade&ID="&Rs(1)&"'>编辑权限</a> <a href=#this onclick=checkclick('"
	If Rs(3)=0 Then Response.Write"注意：等级组必需有一个等级的发帖数为0\n删除可能会导致不能正常！\n\n" 
	Response.Write"删除将会同时调整属于该等级组的用户！\n\n您确定要删除吗？','Admin_Confirm.asp?action=DelGrade&ID="&Rs(1)&"') >"&IconD&"删除"
	Rs.moveNext
	Loop
	Rs.Close
	Response.Write S&"</table><div class='bottom'><input class='button' value='批量更新' type='button'  onclick=""if(confirm('如果你更改了帖数，将会重新修正该组用户的等级！\n如果用户多，会消耗大量资源！\n\n确定更新吗？'))form.submit()"" /><input type='reset' class='button' value='重 置'>&nbsp;&nbsp; &nbsp;&nbsp;<input class='button' value='添加等级' type='button' onclick=window.location.href='?Action=A_E_Grade&Grouping=0' /></div></div></form>"
	Response.Write GoForm("AllUpdateGrade")&"<input name='Grouping' type='hidden' value='1' /><div class='mian'><div class='top'>自定义特别等级组</div><table border='0' class='Stable' cellpadding='3' cellspacing='0'><tr><th width='15%'>等级名称</th><th width='30%'>等级图片</th><th width='19%'>标志图片</th><th width='20%'>管理操作</th></tr>"
	Set Rs=BBS.execute("Select Grouping,ID,GradeName,EssayNum,PIC,Spic,Flag FROM [Grade] where Grouping=1 order by ID")
	do while Not Rs.Eof
	Response.Write"<tr><td align='center'><input name='ID' type='hidden' value='"&Rs(1)&"' />"&_
	"<input class='text' name='GradeName' type='text' size='15' value='"&Rs(2)&"' /></td>"&_
	"<td><input class='text' name='Pic' type='text' size='15' value='"&Rs(4)&"' /><img src='Pic/Grade/"&Rs(4)&"' /></td>"&_
	"<td><input class='text' name='Spic' type='text' size='15' value='"&Rs(5)&"' />"
	If len(Rs(5))>3 Then Response.Write"<img src='Pic/Grade/"&Rs(5)&"' /></td>"
	Response.Write "<td>"&IconE&"<a href='?action=A_E_Grade&ID="&Rs(1)&"'>编辑权限</a> <a href=#this onclick=checkclick('删除将会同时调整属于该等级组的用户！\n\n您确定要删除吗？','Admin_Confirm.asp?action=DelGrade&ID="&Rs(1)&"') >"&IconD&"删除"
	Rs.moveNext
	Loop
	Rs.Close
	Response.Write S&"</table><div class='bottom'><input class='button' value='批量更新' type='submit' /><input type='reset' class='button' value='重 置'>&nbsp;&nbsp; &nbsp;&nbsp;<input class='button' value='添加等级' type='button' onclick=window.location.href='?Action=A_E_Grade&Grouping=1' /></div></div></form>"

	Response.Write GoForm("AllUpdateGrade")&"<input name='Grouping' type='hidden' value='2' /><div class='mian'><div class='top'>系统固定等级组</div><table border='0' class='Stable' cellpadding='3' cellspacing='0'><tr><th width='15%'>等级名称</th><th width='30%'>等级图片</th><th width='19%'>标志图片</th><th width='8%'>属性</th><th width='12%'>管理操作</th></tr>"
	Set Rs=BBS.execute("Select Grouping,ID,GradeName,EssayNum,PIC,Spic,Flag FROM [Grade] where Grouping=2 order by Flag")
	do while Not Rs.Eof
	Response.Write"<tr><td align='center'><input name='ID' type='hidden' value='"&Rs(1)&"' />"&_
	"<input class='text' name='GradeName' type='text' size='15' value='"&Rs(2)&"' /></td>"&_
	"<td><input class='text' name='Pic' type='text' size='15' value='"&Rs(4)&"' /><img src='Pic/Grade/"&Rs(4)&"' />"&_
	"<td><input class='text' name='Spic' type='text' size='15' value='"&Rs(5)&"' />"
	If len(Rs(5))>3 Then Response.Write"<img src='Pic/Grade/"&Rs(5)&"' />"
	If Rs(6)=9 Then Response.Write"</td><td align='center'>站长"
	If Rs(6)=8 Then Response.Write"</td><td align='center'>超版"
	If Rs(6)=7 Then Response.Write"</td><td align='center'>版主"
	If Rs(6)=4 Then Response.Write"</td><td align='center'>VIP"
	Response.Write "</td><td>"&IconE&"<a href='?action=A_E_Grade&ID="&Rs(1)&"'>编辑权限</a> "
	Rs.moveNext
	Loop		
	Rs.Close
	Response.Write S&"</table><div class='bottom'><input class='button' value='批量更新' type='submit' /><input type='reset' class='button' value='重 置'></div></div></form>"
End Sub

Sub A_E_Grade()
	Dim Title,S,Grouping,ID,GradeName,EssayNum,PIC,Spic,Flag,Strings
	ID=request.querystring("ID")
	Grouping=request.querystring("Grouping")
	If ID<>"" Then
		Set Rs=BBS.execute("Select Grouping,ID,GradeName,EssayNum,PIC,Spic,Flag,Strings FROM [Grade] where ID="&ID)
		If Rs.Eof Then
			Goback"","记录不存":Exit Sub
		Else
			Title="编辑等级组"
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
		Title="添加等级组"
		Strings=Split("#F00|1|0|32100|0|1|0|0|1|1|100|1|50|16000|1|1|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0","|")
	End If
	
	If Grouping=1 Then
		Title=Title&"(特别定制)"
	ElseIf Grouping=2 Then
		Title=Title&"(系统固定)"
	Else
		Title=Title&"(按发帖数升级)"
	End If

	Response.Write GoForm("SaveGrade")&"<div class='mian'><div class='top'>"&Title&"</div><input name='ID' type='hidden' value='"&ID&"' /><input name='Grouping' type='hidden' value='"&Grouping&"' />"
	DIVTR"等级名称：","","<input name='GradeName' type='text' class='text' size='15' value='"&GradeName&"' />",25,1
	If Grouping=0 Then DIVTR"必需达到帖数：","","<input name='EssayNum' type='text' class='text' size='4' value='"&EssayNum&"' />",25,1
	If Pic<>"" Then S="<br /><img src='Pic/Grade/"&Pic&"' />" Else S=""
	DIVTR"等级图片：","图片目录\PIC\Grade\","<input name='Pic' type='text' class='text' size='15' value='"&PIC&"' />"&S,42,1
	If sPic<>"" Then S="<img src='Pic/Grade/"&Spic&"' />" Else S=""
	DIVTR"身份标志图片：","图片目录\PIC\Grade\","<input name='Spic' type='text' class='text' size='15' value='"&Spic&"' />"&S,42,1
	Response.Write"<div class='divth'><li><b>基本权限设置</b></li></div><div style='clear: both;'></div>"
	DIVTR"帖子显示名字颜色：","","<input name='S0' class='text' type='text' size='15' maxlength='7' value='"&Strings(0)&"' />",25,1
	DIVTR"是否可以修改自己资料：","",GetRadio("S1","否",Strings(1),0)&GetRadio("S1","是",Strings(1),1),25,2
	DIVTR"是否可以自定义头衔：","",GetRadio("S2","否",Strings(2),0)&GetRadio("S2","是",Strings(2),1),25,1
	DIVTR"帖子最大的字符数：","","<input name='S3' type='text' class='text' size='15' value='"&Strings(3)&"' />个字符(最大不能超过65536)",25,2
	DIVTR"是否可以发表醒目标题：","",GetRadio("S4","否",Strings(4),0)&GetRadio("S4","是",Strings(4),1),25,1
	DIVTR"是否可以参加投票活动：","",GetRadio("S5","否",Strings(5),0)&GetRadio("S5","是",Strings(5),1),25,2
	DIVTR"是否可以发表投票主题：","",GetRadio("S6","否",Strings(6),0)&GetRadio("S6","是",Strings(6),1),25,1
	DIVTR"是否可以上传：","",GetRadio("S8","否",Strings(8),0)&GetRadio("S8","是",Strings(8),1),25,2
	DIVTR"一天的上传个数：","","<input name='S9' type='text' class='text' size='15' value='"&Strings(9)&"' />个",25,1
	DIVTR"每个上传大小：","","<input name='S10' type='text' class='text' size='15' value='"&Strings(10)&"' />KB",25,2
	DIVTR"是否可以上传头像：","",GetRadio("S11","否",Strings(11),0)&GetRadio("S11","是",Strings(11),1),25,1
	DIVTR"论坛信箱最大条数：","","<input name='S12' type='text' class='text' size='15' value='"&Strings(12)&"' />条",25,2
	DIVTR"限制每天发送信件的次数：","","<input name='S7' type='text' class='text' size='15' value='"&Strings(7)&"' />",25,1
	DIVTR"限制每封信字符数：","","<input name='S13' type='text' class='text' size='15' value='"&Strings(13)&"' />个字符(最大不能超过65536)",25,2
	DIVTR"是否可以搜索论坛：","",GetRadio("S14","否",Strings(14),0)&GetRadio("S14","是",Strings(14),1),25,1
	DIVTR"是否可以查看它人信息：","",GetRadio("S15","否",Strings(15),0)&GetRadio("S15","是",Strings(15),1),25,2
	DIVTR"不受时间限制编辑自己帖子：","",GetRadio("S16","否",Strings(16),0)&GetRadio("S16","是",Strings(16),1),25,1
	DIVTR"开启可以删除自己的帖子：","",GetRadio("S17","否",Strings(17),0)&GetRadio("S17","是",Strings(17),1),25,2
	Response.Write"<div class='divth'><li><b>管理权限设置</b> 版主只能管理其管理的版面，其它不限（建议以下选项不要随便开启给按帖数升级的等级组）</li></div><div style='clear: both;'></div>"
	DIVTR"可以编辑帖子：","",GetRadio("S18","否",Strings(18),0)&GetRadio("S18","是",Strings(18),1),25,1
	DIVTR"编辑帖子留下蛛迹的选项：","",GetRadio("S19","否",Strings(19),0)&GetRadio("S19","是",Strings(19),1),25,2
	DIVTR"可以删除帖子：","",GetRadio("S20","否",Strings(20),0)&GetRadio("S20","是",Strings(20),1),25,1
	DIVTR"可以屏蔽帖子：","",GetRadio("S21","否",Strings(21),0)&GetRadio("S21","是",Strings(21),1),25,2
	DIVTR"可以移动帖子：","",GetRadio("S22","否",Strings(22),0)&GetRadio("S22","是",Strings(22),1),25,1
	DIVTR"可以提升主题：","",GetRadio("S23","否",Strings(23),0)&GetRadio("S23","是",Strings(23),1),25,2
	DIVTR"可以沉底主题：","",GetRadio("S24","否",Strings(24),0)&GetRadio("S24","是",Strings(24),1),25,1
	DIVTR"可以(设/解)置顶主题：","",GetRadio("S25","否",Strings(25),0)&GetRadio("S25","是",Strings(25),1),25,2
	DIVTR"可以(设/解)区置顶主题：","",GetRadio("S26","否",Strings(26),0)&GetRadio("S26","是",Strings(26),1),25,1
	DIVTR"可以(设/解)总置顶主题：","",GetRadio("S27","否",Strings(27),0)&GetRadio("S27","是",Strings(27),1),25,2
	DIVTR"可以(设/解)精华主题：","",GetRadio("S28","否",Strings(28),0)&GetRadio("S28","是",Strings(28),1),25,1
	DIVTR"可以(设/解)锁定主题：","",GetRadio("S29","否",Strings(29),0)&GetRadio("S29","是",Strings(29),1),25,2
	DIVTR"可以进行评帖奖罚操作：","",GetRadio("S30","否",Strings(30),0)&GetRadio("S30","是",Strings(30),1),25,1
	DIVTR"可以不需要投票可查投票详情：","",GetRadio("S31","否",Strings(31),0)&GetRadio("S31","是",Strings(31),1),25,2
	DIVTR"可以编辑投票的选项：","",GetRadio("S32","否",Strings(32),0)&GetRadio("S32","是",Strings(32),1),25,1
	DIVTR"可以不受特殊帖限制：","",GetRadio("S33","否",Strings(33),0)&GetRadio("S33","是",Strings(33),1),25,2
	DIVTR"可以发布论坛公告：","",GetRadio("S34","否",Strings(34),0)&GetRadio("S34","是",Strings(34),1),25,1
	DIVTR"可以删除评帖记录：","",GetRadio("S35","否",Strings(35),0)&GetRadio("S35","是",Strings(35),1),25,2
	Response.Write"<div class='divth'><li><b>高级管理权限设置</b> 建议以下选项只给管理员开启</li></div><div style='clear: both;'></div>"
	DIVTR"可以查看用户IP：","",GetRadio("S36","否",Strings(36),0)&GetRadio("S36","是",Strings(36),1),25,1
	DIVTR"可以查看论坛日志：","",GetRadio("S37","否",Strings(37),0)&GetRadio("S37","是",Strings(37),1),25,2
	'DIVTR"开启可以批量操作：","",GetRadio("S38","否",Strings(38),0)&GetRadio("S38","是",Strings(38),1),25,2
	Response.Write "<div class='bottom'><input type='submit' class='button' value='提 交' /><input type='reset' class='button' value='重 置'></div></div></form>"
End Sub

Sub BoardAdmin
	Dim I,po,ii,Name
	Name=Request("Name")
	Response.Write GoForm("BoardAdmin")
	Response.Write"<div class='mian'><div class='top'>增删论坛版主</div>"
	Response.Write"<div class='divtr1' style='padding:3px;'><strong>"&BBS.GetGradeName(0,7)&"</strong>：<input name='Name' type='text' class='text' size='12' value='"&Name&"' /> 操作：<select size='1' name='Flag'><option value='Add'>添加</option><option value='Del'>撤消</option></select> 管理论坛：<select size='1' name='BoardID'><option value=''>请选择管理的版面</option>"&BBS.BoardIDList(0,0)&"</select> <input type='submit' class='button' value='提 交'></div></form>"
	Response.Write GoForm("AllBoardAdmin")
	Response.Write"<div class='divtr2' style='padding:3px;'><strong>"&BBS.GetGradeName(0,8)&"</strong>：<input name='Name' type='text' class='text' size='12' value='"&Name&"' /> 操作：<select size='1' name='Flag'><option value='Add'>添加</option><option value='Del'>撤消</option></select> 管理论坛：全部论坛版面 <input type='submit' class='button' value='提 交'></div></form>"
	Response.Write"</div>"
	Response.Write"<div class='mian'><div class='top'>现有"&BBS.GetGradeName(0,8)&"</div>"
	Response.Write"<div class='divtr1' style='padding:3px;'>"
	Set Rs=BBS.Execute("Select Name From [Admin] where boardID=-1")
	Do while Not Rs.eof
	Po=Po&"<a href='Admin_user.asp?action=EditUser&Name="&Rs(0)&"'>"&Rs(0)&"</a> &nbsp; &nbsp;"
	Rs.movenext
	loop
	Rs.close
	Response.Write po&"</div></div>"
	Response.Write"<div class='mian'><div class='top'>现有"&BBS.GetGradeName(0,7)&"</div>"
	If Not IsArray(BBS.Board_Rs) Then BBS.GetBoardCache()'读取版块缓存
	If IsArray(BBS.Board_Rs) Then
		For i=0 To Ubound(BBS.Board_Rs,2)
			po=""
			For II=1 To BBS.Board_Rs(0,i)
				Po=Po&" <font color=red>∣</Font> "
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
	Response.Write"<div class='mian'><div class='top'><a style='FLOAT: right;color:#FFF' href='?action=A_E_Faction'>"&IconA&"添加帮派&nbsp;</a>论坛帮派管理</div><table class='Stable' border='0' cellpadding='3' cellspacing='0'>"&_
	"<tr><th>帮派名称</th><th>掌门人</th><th>人数</th><th>创派时间</th><th>操作</th></tr>"
	Set Rs=BBS.Execute("Select ID,Name,User,BuildDate From [Faction] order by ID desc")
	Do while Not Rs.eof
	UserNum=BBS.Execute("select count(ID) from [User] where Faction='"&Rs(1)&"'")(0)
	Response.Write"<tr><td>"&Rs(1)&"</td><td align='center'>"&Rs(2)&"</td><td align='center'>"&UserNum&"</td><td align='center'>"&Rs(3)&"</td>"
	Response.Write"<td align='center'><a href='?Action=A_E_Faction&ID="&Rs(0)&"'>"&IconE&" 编辑</a> <a href=#this onClick=""checkclick('删除后将不能恢复！您确定要删除吗？','Admin_Confirm.asp?Action=DelFaction&Name="&Rs(1)&"')"">"&IconD&"删除</a></td></tr>"
	Rs.movenext
	Loop
	Rs.Close
	Response.Write"</table></div>"
End Sub

Sub A_E_Faction
	Dim ID,Name,FullName,Note,User,BuildDate,Title
	Id=Request("ID")
	BuildDate=BBS.NowBbsTime
	Title="添加帮派"
	If ID<>"" Then
		Set Rs=BBS.Execute("Select Name,FullName,Note,User,BuildDate From [Faction] where ID="&ID)
		IF Rs.eof Then Goback"","记录不存在":Exit Sub
		Name=Rs(0)
		FullName=Rs(1)
		Note=Rs(2)
		User=Rs(3)
		BuildDate=Rs(4)
		Title="编辑帮派"
		Rs.Close
	End If
	Response.Write GoForm("SaveFaction")
	Response.Write"<div class='mian'><div class='top'>"&Title&"</div>"
	DIVTR "帮派名称：","","<input name='ID' type='hidden' value='"&ID&"' /><input name='Name' class='text' type='text' size='38' value='"&Name&"' />",22,1
	DIVTR "帮派全称：","","<input name='FullName' type='text' value='"&FullName&"' maxlength='150' size='50' class='text' />",22,2
	DIVTR "帮派宗旨：","","<input name='Note' type='text' class='text' value='"&Note&"' size='50' maxlength='250' />",22,1
	DIVTR "掌门人","","<input type='text' name='User' size='10' class='text' value='"&User&"' /> 用户必须存在",22,2
	DIVTR "创派日期","","<input type='text' name='BuildDate' size='20' class='text' value='"&BuildDate&"' />",22,1
	Response.Write"<div class='bottom'><input type='submit' class='button' value='提 交' /><input class='button' type='reset' value='重 置' /></div></div></form>"
End Sub

Sub DelEssay
Response.Write GoForm("DelEssay&Go=Date")
Response.Write "<div class='mian'><div class='top'>删除指定日期前的帖子</div>"
DIVTR "删除多少天前的帖子：","","<input name='DateNum' type='text' class='text' value='365' size='5'> 天",22,1
DIVTR "选择所在的论坛版面：","","<select name='BoardID'><option value='0'>所有的论坛</option>"&BBS.BoardIDList(0,0)&"</select>",22,1
Response.Write "<div class='divtr2' style='padding:3px;'>说明：此操作将删除指定天数前发表的主题帖，同时也包括主题的回复帖(当然，该主题最新的回复帖也会被删除)。</div>"
Response.Write"<div class='bottom'><input type='button' class='button' value='提 交' onclick=""if(confirm('此操作不可恢复，确定删除吗？'))form.submit()"" /><input class='button' type='reset' value='重 置' /></div></div></form>"
Response.Write GoForm("DelEssay&Go=DateNoRe")
Response.Write "<div class='mian'><div class='top'>删除指定日期前没有回复的主题</div>"
DIVTR "删除多少天前的帖子：","","<input name='DateNum' type='text' class='text' value='100' size='5'> 天",22,1
DIVTR "选择所在的论坛版面：","","<select name='BoardID'><option value='0'>所有的论坛</option>"&BBS.BoardIDList(0,0)&"</select>",22,1
Response.Write "<div class='divtr2' style='padding:3px;'>说明：此操作将删除指定天数前没有再回复主题帖，同时也包括主题的回复帖。</div>"
Response.Write"<div class='bottom'><input type='button' class='button' value='提 交' onclick=""if(confirm('此操作不可恢复，确定删除吗？'))form.submit()"" /><input class='button' type='reset' value='重 置' /></div></div></form>"
Response.Write GoForm("DelEssay&Go=User")
Response.Write "<div class='mian'><div class='top'>删除指定用户的所有帖子</div>"
DIVTR "请输入用户的名称：","","<input name='Name' type='text' value='' class='text' size='20'>",22,1
DIVTR "选择所在的论坛版面：","","<select name='BoardID'><option value='0'>所有的论坛</option>"&BBS.BoardIDList(0,0)&"</select>",22,1
Response.Write "<div class='divtr2' style='padding:3px;'>说明：此操作将删除指定用户的所有帖子。</div>"
Response.Write"<div class='bottom'><input type='button' class='button' value='提 交' onclick=""if(confirm('此操作不可恢复，确定删除吗？'))form.submit()"" /><input class='button' type='reset' value='重 置' /></div></div></form>"
End Sub

Sub DelSms
Response.Write GoForm("DelSms&Go=Date")
Response.Write "<div class='mian'><div class='top'>删除指定日期前的所有留言</div>"
DIVTR "删除多少天前的留言：","","<input name='DateNum' type='text' class='text' value='60' size='5' /> 天",22,1
Response.Write"<div class='bottom'><input type='button' class='button' value='提 交' onclick=""if(confirm('此操作不可恢复，确定删除吗？'))form.submit()"" /><input class='button' type='reset' value='重 置' /></div></div></form>"

Response.Write GoForm("DelSms&Go=Auto")
Response.Write "<div class='mian'><div class='top'>删除自动发送的信件</div>"
DIVTR "删除多少天前自动发送的信件：","","<input name='DateNum' type='text' class='text' value='60' size='5' /> 天",22,1
Response.Write"<div class='bottom'><input type='button' class='button' value='提 交' onclick=""if(confirm('此操作不可恢复，确定删除吗？'))form.submit()"" /><input class='button' type='reset' value='重 置' /></div></div></form>"

Response.Write GoForm("DelSms&Go=User")
Response.Write "<div class='mian'><div class='top'>删除指定用户的所有留言</div>"
DIVTR "请输入指定用户名称：","","<input name='Name' type='text' class='text' value='' size='20' />",22,1
Response.Write"<div class='bottom'><input type='button' class='button' value='提 交' onclick=""if(confirm('此操作不可恢复，确定删除吗？'))form.submit()"" /><input class='button' type='reset' value='重 置' /></div></div></form>"
End Sub

Sub MoveEssay
	Response.Write GoForm("MoveEssay&Go=Date")
	Response.Write "<div class='mian'><div class='top'>按指定天数移动帖子</div>"
	DIVTR"移动多少天前的帖子：","","<input name='DateNum' type='text' class='text' value='100' size='5' /> 天",22,2
	DIVTR"帖子原来所在的论坛：","","<select size='1' name='BoardID1'>"&BBS.BoardIDList(0,0)&"</select>",22,1
	DIVTR"帖子要移动到的论坛：","","<select size='1' name='BoardID2'>"&BBS.BoardIDList(0,0)&"</select>",22,1
	Response.Write"<div class='bottom'><input type='button' class='button' value='提 交' onclick=""if(confirm('您确定要移动帖子吗？'))form.submit()"" /><input class='button' type='reset' value='重 置' /></div></div></form>"
	Response.Write GoForm("MoveEssay&Go=User")
	Response.Write "<div class='mian'><div class='top'>移动指定用户的帖子</div>"
	DIVTR"请输入指定的用户名：","","<input name='Name' type='text'  size='20' class='text' />",22,2
	DIVTR"帖子原来所在的论坛：","","<select size='1' name='BoardID1'>"&BBS.BoardIDList(0,0)&"</select>",22,1
	DIVTR"帖子要移动到的论坛：","","<select size='1' name='BoardID2'>"&BBS.BoardIDList(0,0)&"</select>",22,1
	Response.Write"<div class='bottom'><input type='button' class='button' value='提 交' onclick=""if(confirm('您确定要移动这个用户的帖子吗？'))form.submit()"" /><input class='button' type='reset' value='重 置' /></div></div></form>"
End Sub

Sub TopAdmin
	Dim Flag,bgcolor,I,S
	If Instr(AdminString,",22,")=0 Then
	Showtable"后台权限","<li>你被禁止操作管理员的权限！！</li><li><a href='Admin_user.asp?Action=AdminOK&Name="&BBS.MyName&"'>只能修改自己的密码</a></li>"
	Footer()
	Response.End
	End If
	Response.Write "<form method=POST  name=form style='margin:0' action='Admin_Confirm.asp?Action=TopAdmin&Flag=1'>"
	Response.Write "<div class='mian'><div class='top'>添加论坛管理员</div>"
	DIVTR"用户名称：","","<input name='Name' type='text' class='text' size='20'>",22,1
	Response.Write"<div class='bottom'><input type='submit' class='button' value='提 交' /><input class='button' type='reset' value='重 置' /></div></div></form>"
	Response.Write "<div class='mian'><div class='top'>"&BBS.GetGradeName(0,9)&"列表</div>"
	i=0
	Set Rs=BBS.execute("Select Name,BoardID From[Admin] where BoardID=0")
	Do while not Rs.eof
	S="<a href=#this onclick=""checkclick('您确定要取消其管理员的职位吗？','Admin_Confirm.asp?action=TopAdmin&name="&Rs(0)&"&Flag=0')"">【降职】</a>"
	IF Rs(0)=BBS.GetMemor("Admin","AdminName") Then S="<a onclick=alert('不能对自己降职！') href='#'>【降职】</a>"
	S=S&"<a href='Admin_user.asp?action=AdminOK&Name="&Rs(0)&"'>【设置后台权限】</a>"
	DIVTR "<a href='admin_User.asp?action=EditUser&Name="&Rs(0)&"'>"&Rs(0)&"</a>","","<div style='padding:3px'>"&S&"</div>",22,2
	Rs.movenext
	Loop
	Rs.Close
	Response.Write"</div>"
End Sub

Sub Clean
	Response.Write GoForm("Clean")
	Response.Write "<div class='mian'><div class='top'>更新空间缓存</div>"
	Response.Write "<div class='divtr2' style='padding:3px;'>论坛采用了服务器缓存技术，使论坛的速度飞快，如果发现论坛出现不稳定的状态，可以清空论坛的缓存。<br />论坛缓存采用了实时更新设计，一般情况下不建议更新论坛的缓存。<br />在线人员信息采用缓存记录，更新将全部清除！</div>"
	Response.Write "<div class='bottom'><input type='submit' class='button' value='更新本站全部缓存'></div></div></form>"
End Sub

Sub Bank
	Response.Write"<form method=POST  style='margin:0' action='Admin_Confirm.asp?Action=Bank' onSubmit=""ok.disabled=true;ok.value='银行正在处理-请稍等。。。'"">"
	Response.Write "<div class='mian'><div class='top'>后台银行</div>"
	DIVTR"用户群：","","<select name='user' style='font-size: 9pt'><option value='' selected></option><option value=1>所有在线用户</option><option value=7>"&BBS.GetGradeName(0,7)&"</option><option value=8>"&BBS.GetGradeName(0,8)&"</option><option value=9>"&BBS.GetGradeName(0,9)&"</option><option value=10>管理团队(版主+管理员)</option><option value=4>所有Vip用户</option><option value=0>所有注册用户(慎用)</option></select>",25,1
	DIVTR"操作：","","<input name='Flag' type='radio' value='1' checked>增加 <input name='Flag'  type='radio' value='0'>减少",25,1
	DIVTR BBS.Info(120)&"：","","<input name='Coin' type='text' value='1000' >",25,1
	Response.Write "<div class='bottom'><input  type='submit' class='button' value='确定' name='ok'></div></div></form>"
End Sub


Sub GapAd
	Response.Write GoForm("GapAd")
	Response.Write "<div class='mian'><div class='top'>贴间广告管理</div>"
	Response.Write "<div class='divth' style='padding:5px;text-align: left;'>说明：这些文字广告将会在帖子与帖子之间随机显示。<br />请使用简单的文字超连接html代码。<br />将代码清空则删除相关广告。</div>"
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
	 	DIVTR I&"、显示效果：","","<div style='line-height:25px'>"&ad_tmp&"</div>",25,2
		DIVTR "&nbsp;&nbsp;&nbsp;相应代码：","","<textarea  rows='3'  name=ad_v"&i&">"&ad_tmp&"</textarea>",50,1
	Next
	DIVTR"<span style='color:#F00'>增加广告：</span>","","<textarea  rows='3'  name=ad_v"&ad_num+1&"></textarea>",50,2
	Response.Write"<div class='bottom'><input type='submit' class='button' value='确定修改' /><input class='button' type='reset' value='重 置' /></div></div></form>"
	OpenFile.close
	Set FSO=Nothing
End Sub
%>
