<!--#include file="inc.asp"-->
<%
Dim ErrNum,NeedLogin,GoBack,Caption,Content
NeedLogin=False
GoBack=False
BBS.Head "err.asp?errnum="&request.querystring("ErrNum"),"","错误信息"
Caption="错误信息！"
ErrNum=request.querystring("ErrNum")
If Not isnumeric(ErrNum) Then ErrNum=1
Select Case ErrNum
	Case 1
		Caption="非法操作！"
		Content = "<li>错误的地址栏参数，请不要手动去更改地址栏参数。</li>"
	Case 2
		Caption="非法操作 ！"
		Content = "<li>您所提交的数据来自论坛外部，请不要从论坛外部提交数据，谢谢合作！</li>"
	Case 3
		NeedLogin = True
		Content ="<li>必须登录论坛后才能进行操作！请<a href='login.asp'>登陆</a>！</li>"
	Case 4
		Caption="非法操作 ！"
		Content ="<li>请不要用软件提交！</li>"
	Case 5
		GoBack = True
		Caption="登陆失败"
		Content="<li>本站为了防止恶意尝试机器登陆，2次登陆间隔被设为<Font color=red>"&BBS.Info(10)&"</Font>分钟</li>"
	Case 6
		GoBack=True
		NeedLogin = True
		Caption ="登陆失败"
		Content = "<li>亲爱的用户，请你在登陆时不要忘了填用户名或密码。</a></li>"
	Case 7
		GoBack=True
		Caption ="登陆失败"
		Content = "<li>请填写验证码。</li>"	
	Case 8
		GoBack=True
		NeedLogin = True
		Content = "<li>请填写正确的验证码</li>"	
	Case 9
        NeedLogin = True
		GoBack = True
		Caption="登陆失败"
		Content="<li>您的用户名或密码错误</li><li>或此账号被暂时删除！</li><li>注意：如果错误登陆超过5次，您今天将不能用此帐号登陆！</li>"
	Case 10
		GoBack=True
		NeedLogin = True
		Caption="进入失败 ！"
		Content="<li>您不能成功的进入该版面！</li><li>该版面为只有注册会员可以进入！</li><li>你还没有<a href=login.asp>登陆</a>！</li>"
	Case 11
		GoBack=True
		Caption="进入失败 ！"
		Content="<li>您不能成功的进入该版面！</li><li>该版面不存在！</li><li> 请正确访问本论坛，现在回到<a href=Index.asp>首页</a>！</li>"
	Case 12
		GoBack=True
		Caption="进入失败 ！"
		Content="<li>您不能成功的进入该版面！</li><li>该版面已被锁定！停止开放!</li>"
	Case 13
		GoBack=True
	   	Caption="进入失败 ！"
		Content="<li>你不能成功的进入该版面！</li><li> 该版面为VIP论坛，您还不是VIP用户！</li>"		
	Case 14
		GoBack=True
	   	Caption="进入失败 ！"
		Content="<li>你不能成功的进入该版面！</li><li> 该版面为认证论坛，你还没有得版主的认证！</li>"
	Case 15
		GoBack=True
	   	Caption="进入失败 ！"
		Content="<li>你不能成功的进入该版面！</li><li> 本版面为限制等级,您目前的论坛等级达不到该版面的要求！</li>"
	Case 16
		GoBack=True
	   	Caption="进入失败 ！"
		Content="<li>你不能成功的进入该版面！</li><li> 本版面为限制发帖数,您目前的发帖数量达不到该版面的要求！</li>"
	Case 17
		GoBack=True
	   	Caption="进入失败 ！"
		Content="<li>你不能成功的进入该版面！</li><li> 本版面为限制积分,您目前的积分达不到该版面的要求！"				
	Case 18
		GoBack=True
	   	Caption="进入失败 ！"
		Content="<li>你不能成功的进入该版面！</li><li> 本版面为限制金钱,您目前的金钱数量达不到该版面的要求！</li>"
	Case 19
		GoBack=True
	   	Caption="进入失败 ！"
		Content="<li>你不能成功的进入该版面！</li><li> 本版面为限制游戏币,您目前的游戏币数量达不到该版面的要求！</li>"				 
	Case 20
        NeedLogin = True
		GoBack=True
		Caption="用户操作"
		Content="<li>你不能成功的进入该页面！</li><li>该页面为只有注册会员可以进入！</li><li>你还没有<a href='login.asp'>登陆</a></li>！"
	Case 21
		GoBack=True
		Content="<li>你的操作出错！</li><li>该帖子不存在</li><li>或该帖子已经删除</li><li>返回<a href='Index.asp'>论坛首页</a></li>" 
	Case 22
		GoBack = True
		Content="<li>你的操作出错！</li><li>该帖子已经被锁定！</li><li><a href='Index.asp'>返回论坛首页</a></li>"
	Case 23
		Caption="注册失败 ！"
		GoBack = True
		Content = "<li>抱歉，论坛暂停新用户注册！</li>" 
	Case 24
		Caption="注册失败 ！"
		GoBack=True
		Content="<meta http-equiv=refresh content=4;url=Index.asp><li>对不起！你不能成功注册！！！</li><li>本论坛为了防止机器注册等恶意注册，同一用户来源限制注册间隔<b>"&BBS.Info(9)&"</b> 分钟！</li>"
	Case 25
        NeedLogin = True
		GoBack=True
		Caption="进入失败"
		Content="<li>你没有浏览精华帖子的权限</li><li>只有注册会员才可以进入！</li><li> 你还没有<a href=""login.asp""> 登陆 </a>！</li>"
    Case 26
        NeedLogin = True
		GoBack = True
		Content="<li>你目前的不是会员，只有游览开放帖子的权限</li>"
	Case 27
		GoBack = True
		Content = "<li>对不起！你不能成功地发出帖子！</li><li> 你并没有填写标题或内容！</li>"
	Case 28
		GoBack = True
		Content = "<li>对不起！你不能成功地发出帖子！</li><li>帖子标题字符数超出论坛限制！</li>" 
    Case 29
		GoBack = True
		Content = "<li>内容字符数超出论坛限制！</li>"
 	Case 30
		GoBack=True
		Content = "<li>对不起！你不能成功地发出帖子！</li><li>验证码不对，请填写正确的验证码</li>"	
	Case 31
	    NeedLogin = True
		GoBack=True
		Content = "<li>你还没有 <a href=""login.asp"">登陆</a> 或 <a href=""register.asp"">注册</a> ！</li>"	
 	Case 32
		GoBack=True
		Content="<li>你的操作出错！</li><li>该帖子不存在 或 已经删除</li><li>返回<a href=Index.asp>论坛首页</a></li>"
	Case 33 
		GoBack = True
		Content ="<li>你的操作出错！</li><li>你不是该帖的作者或该版面的斑竹，所以不能编辑该帖子</li>"
	Case 34
		GoBack = True
		Content = "<li>你不能编辑帖子！</li><li>因为你超过了允许普通用户编辑自己帖子的时限 (即发帖后<font color='#F00'>"&BBS.Info(12)&"</font>分钟内)</li>"
	Case 35
		GoBack=True
		Content = "<li>帖子发送失败！</li><li>验证码失效，请填写正确的验证码</li>"	
	Case 36
		GoBack = True
		Content = "<li>您还没有填写完整的必填选项</li>"
	Case 37
		GoBack = True
		Content = "<li>请不要使用非法字符或者属于禁止注册之列！！</li>"		
	Case 38
		GoBack = True
		Content = "<li>用户名 和 密码 必须小于14个字符(7个汉字)或者不能用单个字符！</li>"		
	Case 39
		Caption="注册失败！"
		GoBack = True
		Content = "<li>注册失败，可能是您申请注册的这个昵称已经被另一个用户使用了！</li>"		
	Case 40
		GoBack = True
		Content = "<li>对不起，您的密码属于弱口令，请更改密码!</li>"	
	Case 41
		GoBack = True
		Content = "<li>您两次输入的密码不相同！</li>"		
	Case 42
		GoBack = True
		Content = "<li>请填写正确、有效的EMail地址！</li>"
	Case 43
		GoBack = True
		Content = "<li>“密码问题”或“问题答案”的字符太短，必须大于4个字符！</li>"
	Case 44
		GoBack = True
		Content ="<li>“密码问题”或“问题答案”的含有非法字符！</li>"					   
	Case 45 
		GoBack = True
		Content ="<li>本站设置不允许使用外部头像！</li>"		
	Case 46
		GoBack =True
		Content ="<li>由于你选用QQ形象作为头像，请正确填写你的QQ号码！</li>"		
	Case 47
		GoBack = True
		Content ="<li>您填写的一些项目的字符数超过了论坛的限制！</li>"		
	Case 48
		GoBack = True
		Content = "<li>头像宽度和高度必须用数字填写！</li>"		
	Case 49
		GoBack = True
		Content = "<li>您填写的邮箱已被注册！</li>"	
	Case 50
		GoBack = True
		Content = "<li>经检测，发出的是没有文字的内容！</li>"
	Case 51	
		GoBack = True
		Content = "<li>此帖的不是在本版发表的，所以您不能编辑该帖子</li>"
	Case 52
        GoBack = True
		Content = "<li>对不起，不能发送留言，你的"&BBS.Info(120)&"达不到<font color=red>"&BBS.Info(123)&"</font>！！</li>"
	Case 53
        GoBack = True
		Content = "<li>对不起，本站设定留言间隔1分钟</li>"
	Case 54
        GoBack = True
		Content = "<li>发送失败，论坛中不存在该留言对象</li>"
	Case 55
        GoBack = True
		Content = "<li>发送失败，今天已经不能发送留言了</li>"		
 	Case 56
        GoBack =True
		Content = "<li>你填写的旧密码不正确！</li>"
	Case 57
		GoBack = True
		Content = "<li>找回密码失败！</li><li>您的密码提示问题和问题答案不正确！</li>"
	Case 58
		GoBack=True
		Content="<li>该帖子不存在 或者 已经删除</li><li><a href=""Index.asp"">返回论坛首页</a></li>"
	Case 59
		GoBack = True
		Content = "<li>该帖子已经是总置顶帖！</li>"		
	Case 60
        GoBack = True
		Content = "<li>该帖子已经是区置顶帖！</li>"	
	Case 61
        GoBack = True
		Content = "<li>该主题帖子已经取消了总置顶！</li>"
	Case 62
        GoBack = True
		Content = "<li>你选择的是同一个版面，不用移动了！</li>"	
	Case 63
        GoBack = True
		Content = "<li>搜索的关键字字符长度小于论坛限制的 2 个字符 </li>"
	Case 64
        GoBack = True
		Content = "<li>本论坛限制了每次搜索时间间隔为 "&BBS.Info(16)&" 秒</li>"
	Case 65
        GoBack = True
		Content = "<li>对不起，论坛限制您每天只能可以进行"&BBS.Info(49)&"次评帖操作</li>"
	Case 66
		GoBack = True
		Content = "<li>一些选项必需用数字填定。</li>"
	Case 67
		GoBack = True
		Content = "<li>你所在的等级组不能发布和管理公告！查看 <a href='help.asp?action=mygrade'>我的权限</a></li>"
	Case 68
		GoBack = True
		Content = "<li>你不是该版面的版主！</li>"
	Case 69
		GoBack = True
		Content = "<li>找不到相应的记录，可能已经删除了。</li>"											
	Case 70
        GoBack = True
		Content = "<li>你不能进行操作，请确认你的等级权限！查看 <a href='help.asp?action=mygrade'>我的权限</a></li>"					
	Case 71
        GoBack = True
		Content = "<li>你不能进行操作，你不是该版面的版主！</li>"
	Case 72
		Content = "<li>天啊，你被管理员踢出了论坛！</li><li>在"&BBS.Info(8)&"分钟内你不能登陆论坛！</li>"
	Case 73
		GoBack = True
		Content = "<li>你不能编辑不是你发布的公告！</li>"
	Case 74
        GoBack = True
		Content = "<li>你所在的等级组不能查看用户信息，请确认你的等级权限！查看 <a href='help.asp?action=mygrade'>我的权限</a></li>"
	Case 75
        GoBack = True
		Content = "<li>你所在的等级组不能进行论坛搜索，请确认你的等级权限！查看 <a href='help.asp?action=mygrade'>我的权限</a></li>"
 	Case 76
		GoBack=True
		Content="<li>你不能进行操作，你不是管理员或版主！！</li>"
 	Case 77
		GoBack=True
		Content="<li>对不起！本站没有开放论坛搜索功能!</li>"
	Case 78
		GoBack=True
		Caption="登陆失败！"
		Content="<li>您注册的信息，还没有通过管理员的审核。</li><li>请耐心等待管理员的审核，谢谢合作！</li>"
	Case 79
		GoBack=True
		Caption="找不到用户！"
		Content="<li>该用户资料可能已经删除或者未被审核通过！</li>"
	Case 81
		GoBack = True
		Content = "<li>头像图片的路径含有非法字符！</li>"
	Case 82
		GoBack = True
		Content = "<li>此论坛为限制论坛，只允许版主及以上的会员发帖！</li>"
	Case Else
       Content = "晕~~~为什么^_^呵呵"
End Select
	If GoBack Then Content=Content&"<li><a href=javascript:history.go(-1)>返回上一页</a>"
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
		Temp=Temp&BBS.Row("<b>请输入您的用户名：</b>","<input name=""name"" type=""text"" class=""submit"" size=""20"" /> <a href=""register.asp"">没有注册？</a>","65%","")
		Temp=Temp&BBS.Row("<b>请输入您的密码：</b>","<input name=""Password"" type=""password"" size=""20"" /> <a href=""usersetup.asp?action=forgetpassword"">忘记密码？</a>","65%","")
		If BBS.Info(14)="1" Then
			Temp=Temp&BBS.Row("<b>请输入右边的验证码：</b>",BBS.GetiCode,"65%","")
		Else
			Temp=Temp&"<input name=""iCode"" type=""hidden"" value=""BBS"" />"
		End If
		Temp=Temp&BBS.Row("<b>Cookie 选项：</b>","<input type=radio  name=cookies value=""0"" checked class=checkbox />不保存 <input type=radio  name=cookies value=""1"" class=checkbox />保存一天 <input type=radio  name=cookies value=""30"" class=checkbox />保存一月","65%","")
		Temp=Temp&BBS.Row("<b>选择登陆方式：</b>","<input type=radio value=""1"" checked name='hidden' class=checkbox />正常登陆 <input type='radio' value='2' name='hidden' class=checkbox />隐身登陆","65%","")
		Temp=Temp&"<div style="" padding:5px;BACKGROUND: "&BBS.SkinsPIC(1)&";"" align=""center""><input name=""submit"" type=""submit"" value=""登 陆"" /></div></form>"
		BBS.ShowTable"用户登陆",Temp
	End If
BBS.Footer()
Set BBS =Nothing
%>