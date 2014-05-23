<%
'################################################################################
'## ebody.asp
'## -----------------------------------------------------------------------------
'## 功能:	Ebody核心基本类
'## 版本:	1.0
'## 作者:	Tony
'## 日期:	2014/05/23
'## 说明:	系统的基本类,用于存放通用性或系统必须的底层方法
'################################################################################
'## 系统核心类及扩展类架构
'## -----------------------------------------------------------------------------
'##	ebody
'##		|---tpl(core)
'##		|	|---methods
'##		|---fso(core)
'##		|	|---methods
'##		|---db(core)
'##		|	|---methods
'##		|---upload(core)
'##		|	|---methods
'##		|---json(core)
'##		|	|---methods
'##		|---http(core)
'##		|	|---methods
'##		|---xml(core)
'##		|	|---methods
'##		|---aes(core)
'##		|	|---methods
'##		|---md5(core)
'##		|	|---methods
'##		|---cache(core)
'##		|	|---methods
'##		|---list(core)
'##		|	|---methods
'##		|---errmsg(core)
'##		|	|---methods
'##		|---ext(core)
'##			|---ext1(plugin 1)
'##			|	|---methods
'##			|---ext2(plugin 2)
'##			|	|---methods
'##			|---...
'##			|	|---methods
'##			|---ext(n)(plugin n)
'##				|---methods
'################################################################################
'## 动态加载类及使用方法
'## -----------------------------------------------------------------------------
'## 核心类使用Use加载,扩展类及插件使用Extend加载.
'## 加载tpl核心基类: ebody.use "tpl"

'## 加载music扩展类: ebody.extend "music"

'## 然后后可直接在基类名后引用,而扩展类则需要引用基类ext

'## 如加载核心类tpl后的使用方法为: ebody.tpl.file

'## 如加载扩展类music后的使用方法为: ebody.ext.music.play
'################################################################################
'## 开发者说明
'## -----------------------------------------------------------------------------
'## 1. 如增加核心类功能，需要定义一个同名的全局变量
'## 
'## 
'################################################################################


'################################################################################
'## 全局初始化
'################################################################################

' 定义全局对像
Dim Ebody : Set Ebody = New Ebody_Base

'################################################################################
'## 类主体
'################################################################################

Class Ebody_Base
	
'================================================================================
'== Variable
'================================================================================


'--------------------------------------------------------------------------------
'-- Public
'--------------------------------------------------------------------------------

	' 定义基本类
	Public tpl					' 模板类
	Public fso					' 文件操作类
	Public db					' 数据操作类
	Public upload				' 上传类
	Public json					' JSON类
	Public http					' HTTP远程访问类
	Public xml					' XML类
	Public aes					' AES加密类(easp2.2)
	Public md5					' MD5一致性效验类(easp2.2)
	Public cache				' 缓存类(easp2.2)
	Public list					' 列表数组类(easp2.2)
	Public errmsg				' 错误提示类(easp2.2)
	Public ext					' 扩展类

'--------------------------------------------------------------------------------
'-- Private
'--------------------------------------------------------------------------------

	' 定义类内部变量
	Private cvGlobalClassName	' 系统类名	
	Private cvCoreClassPath		' 核心类库目录
	Private cvExtClassPath		' 扩展类库目录
	Private cvCharset			' 字符集
	Private coFSO				' 文件操作对像
	Private coRegExp			' 正则表达对像
	Private cvErrCode			' 错误代码
	Private cvErrMsg			' 错误信息
	Private cvFsoName			' FSO类名
	Private cvDicName			' 数据字典类名
	Private cvBom				' BOM处理方式
	Private cvCookieEncode		' Cookies是否加密
	Private cvDBConnStr			' DB连接字串

	' 局部通用变量
	Private cvSysTime			' 系统当前时间

'================================================================================
'== Event
'================================================================================

	' 初始化类
	Private Sub Class_Initialize()
		' 设置忽略错误
		'On Error Resume Next
		' 初始通用变量值
		cvSysTime = Timer()							' 当前时间
		' 定义初始属性		
		cvGlobalClassName = "ebody"					' 全局类名(表示系统类名以什么开头)			
		cvCharset = "utf-8"							' 设定文件字符集
		cvBom = "remove"							' 页面bom处理方式
		cvCookieEncode = False						' Cookies是否加密
		cvFsoName = "Scripting.FileSystemObject"	' 全局文件对像类名
		cvDicName = "Scripting.Dictionary"			' 全局数据字典类名		
		' 设置页面编码
		Session.CodePage = GetCharCode(cvCharset)
		Response.Charset = cvCharset
		Err.Clear
		' 取消错误捕捉
		On Error Goto 0
		' 创建ebody类中必要的对象
		Set coFSO = Server.CreateObject(cvFsoName)	' 创建文件系统对像实例
		Set coRegExp = New RegExp					' 创建正则表达对像实例
		coRegExp.Global = True						' 正则查询属性, True: 全局搜索, False: 只搜索一行
		coRegExp.IgnoreCase = True					' 正则匹配属性, True: 不区分大小写, False: 区分大小写
		' 设定基类的物理路径
		cvCoreClassPath = IIF(IsFolder(AbsPath_(cvCoreClassPath)), AbsPath_(cvCoreClassPath), VRoot & Replace(cvCoreClassPath, "/", "\"))
		cvExtClassPath = IIF(IsFolder(AbsPath_(cvExtClassPath)), AbsPath_(cvExtClassPath), VRoot & Replace(cvExtClassPath, "/", "\"))
	End Sub

	' 结束类
	Private Sub Class_Terminate()
		' 设置忽略错误
		'On Error Resume Next
		' 关闭所有Ebody对象
		Call CloseAll
		' 销毁文件系统对像
		If IsObject(coFSO) Then Set coFSO = Nothing
		' 销毁正则表达对像
		If IsObject(coRegExp) Then Set coRegExp = Nothing
		' 清除错误信息
		cvErrCode = empty
		cvErrMsg = empty
	End Sub	

'================================================================================
'== Public
'================================================================================


'--------------------------------------------------------------------------------
'-- Let Property
'--------------------------------------------------------------------------------

	' 设定DB连接字串
	Public Property Let DBConnStr(ByVal pStr)
		cvDBConnStr = pStr
	End Property


	' 设定Ebody类文件存放的基本路径
	' 说明: 当类目录放在网站的多级子目录下，此时就需要指定存放的目录级
	'		比如：网站目录为d:\abc，而ebody类目录放在d:\abc\efg\def\下面，那么这里就需要指定BaesPath为：\efg\def
	Public Property Let BasePath(ByVal pStr)
		cvCoreClassPath = AbsPath_(pStr & "/core")		' 核心类文件所在绝对目录
		cvExtClassPath = AbsPath_(pStr & "/plugin")		' 扩展类文件所在绝对目录
	End Property


	' 设置FSO组件名称
	Public Property Let FsoName(ByVal pStr)
		cvFsoName = pStr
	End Property
	
	' 设定页面字符集
	' 调用: Ebody.CharSet = "UTF-8"
	' 说明: 设定当前页面所使用的字符集
	Public Property Let [CharSet](ByVal pStr)
		cvCharset = Ucase(pStr)
		On Error Resume Next
		Session.CodePage = GetCharCode(cvCharset)
		Response.Charset = cvCharset
		Err.Clear
		On Error Goto 0
	End Property

	' 设置如何处理UTF-8文件的BOM信息
	' 调用: Ebody.FileBOM = "keep"
	' 说明: 处理类型分为 keep/remove/add
	Public Property Let FileBOM(ByVal pStr)
		cvBom = Lcase(pStr)
	End Property

	' 设置是否加密Cookies信息
	' 调用: Ebody.CookieEncode = true
	Public Property Let CookieEncode(ByVal pBool)
		cvCookieEncode = pBool
	End Property	

'--------------------------------------------------------------------------------
'-- Get Property
'--------------------------------------------------------------------------------

	' 取DB连接字串
	Public Property Get DBConnStr()
		DBConnStr = cvDBConnStr
	End Property

	' 读取FSO组件名称
	Public Property Get FsoName()
		FsoName = cvFsoName
	End Property
	
	' 取得服务器验证类型
	' 调用: Ebody.GetHttpType
	' 返回: http 一般验证, https 加密验证
	Public Property Get HttpType()
		Dim lvHttpType
		If LCase(Request.ServerVariables("HTTPS"))="off" Then
			lvHttpType = "http"
		Else
			lvHttpType = "https"
		End If
		HttpType = LCase(lvHttpType)
	End Property

	' 取得服务器端口号
	' 调用: Ebody.PortNo
	' 返回: 8088
	Public Property Get PortNo()
		PortNo = Request.ServerVariables("SERVER_PORT")
	End Property

	' 取得网站根物理路径
	' 调用: Ebody.Root
	' 返回: C:\taoya
	Public Property Get Root()
		Root = Server.MapPath("/")
	End Property

	' 取得网站主页网址
	' 调用: Ebody.Home
	' 返回: http://www.taoya.com
	Public Property Get Home()
		Home = HttpType & "://" & Request.ServerVariables("SERVER_NAME")
	End Property

	' 自动获取当前页面所在的根物理路径
	' 调用: Ebody.VRoot
	' 返回: 根目录与虚拟目录路径一致, 说明文件是在根目录下, 返回根目录, 否则文件在虚拟目录下, 返回虚拟目录
	' 说明: 用户浏览http://localhost/tt/tony/demo/home.asp, 其文件放在一个网站的虚拟目录tt下, 网站在D:盘, 但虚拟目录对应的文件在E:盘
	'		此时就需要使用此方法自动判断网站根地址
	Public Property Get VRoot()
		' 判断根目录是否与虚拟目录的路径一致, 来返回网站根路径
		VRoot = IIF(Root=VDirPath, Root, VDirPath)
	End Property

	' 自动获取当前页面所在的根网址
	' 调用: Ebody.VHome
	' 返回: 页面存放分为 网站根目录与虚拟目录, 当浏览的页面是在根目录下, 返回网站主页, 文件在虚拟目录下, 则虚拟目录的地址作为主页返回
	' 说明: 用户浏览http://localhost/tt/tony/demo/home.asp, 其文件属于网站的虚拟目录tt下, 此时返回http://localhost/tt
	Public Property Get VHome()
		VHome = Home & IIF(VDirName<>"ROOT", "/" & VDirName, "")
	End Property

	' 取得当前页面所在的虚拟目录名称
	' 调用: Ebody.VDirName
	' 返回: 假设当前页面地址http://localhost/tt/tony/demo/home.asp, 虚拟目录是tt, 则返回tt
	Public Property Get VDirName()
		Dim lvAppPath : lvAppPath = Request.ServerVariables("Appl_MD_Path")
		DIM lvInstancePath : lvInstancePath = Request.ServerVariables("Instance_Meta_Path") & "/"
		Dim S : S = Mid(lvAppPath, Len(lvInstancePath)+1)
		If InStr(S, "/")>0 Then S = Mid(S, InStrRev(S, "/")+1)
		VDirName = S
	End Property

	' 取得当前页面所在的虚拟目录的物理路径
	' 调用: Ebody.VDirPath
	' 返回: 假设网站根物理路径是C:\taoya, 虚拟目录是E:\demo, 则返回E:\demo
	Public Property Get VDirPath()
		Dim lvPath
		lvPath = Request.ServerVariables("APPL_PHYSICAL_PATH")
		VDirPath = IIF(Right(lvPath, 1)="\", Left(lvPath, InStrRev(lvPath, "\") - 1), lvPath)
	End Property

	' 20130609
	' 取得服务器端执行时间(秒)
	' 调用: Ebody.ScriptTime
	' 返回: 服务器端的执行时间长度(保留3位小数)
	Public Property Get ScriptTime
		ScriptTime = FormatNumber(GetScriptTime(0)/1000, 3, -1)
	End Property

	' 取得页面字符集
	Public Property Get CharSet()
		CharSet = cvCharset
		'CharSet = Session.CodePage	' 代码
		'CharSet = Response.Charset	' 名称
	End Property

	' 取得如何处理UTF-8文件的BOM信息类型
	Public Property Get FileBOM()
		FileBOM = cvBom
	End Property

	' 取得是否加密Cookies信息
	Public Property Get CookieEncode()
		CookieEncode = cvCookieEncode
	End Property

	' 取得系统错误信息
	Public Property Get [Error]()
		[Error] = cvErrMsg
	End Property

'--------------------------------------------------------------------------------
'-- Function
'--------------------------------------------------------------------------------

	'-------------------------
	' 基础功能类函数
	'-------------------------

	' 判断三元表达式 (来自Easp2.2)
	' 调用: IIF(条件表达式, 成功时返回值, 失败时返回值)
	' 样例: Ebody.IIF(1=2, "YES", "NO")
	' 返回: NO
	Public Function IIF(ByVal pC, ByVal pT, ByVal pF)
		If pC Then
			IIF = pT
		Else
			IIF = pF
		End If
	End Function

	' 如果条件成立则返回某值 (来自Easp2.2)
	Public Function IfThen(ByVal Cn, ByVal T)
		IfThen = IIF(Cn,T,"")
	End Function

	' 如果第1项不为空则返回第1项, 否则返回第2项 (来自Easp2.2)
	' 说明: 相当于oracle中的nvl函数
	Public Function IfHas(ByVal v1, ByVal v2)
		IfHas = IIF(Has(v1), v1, v2)
	End Function

	' 判断是否为空值
	' 调用: IsNull(对像)
	' 样例: IsNull("abc")
	' 返回: False
	Public Function IsNull(ByVal pObj)
		IsNull = False
		Select Case VarType(pObj)
		Case vbEmpty, vbNull
			IsNull = True : Exit Function
		Case vbString
			If pObj="" Then IsNull = True : Exit Function
		Case vbObject
			Select Case TypeName(pObj)
			Case "Nothing","Empty"
				IsNull = True : Exit Function
			Case "Recordset"
				If pObj.State = 0 Then IsNull = True : Exit Function
				If pObj.Bof And pObj.Eof Then IsNull = True : Exit Function
			Case "Dictionary"
				If pObj.Count = 0 Then IsNull = True : Exit Function
			End Select
		Case vbArray,8194,8204,8209
			If Ubound(pObj)=-1 Then IsNull = True : Exit Function
		End Select
	End Function

	' 判断是否有值
	' 调用: Has(对像)
	' 样例: Has("abc")
	' 返回: True
	Public Function Has(ByVal pObj)
		Has = Not IsNull(pObj)
	End Function

	' 不够长度的字符串填充内容 (来自Easp2.2)
	' 样例: Fill("abc", 10, "x"), 返回: abcxxxxxxx
	Public Function Fill(ByVal s, ByVal l, ByVal f)
		Dim le,i,tmp
		le = len(s)
		If isNull(f) Then f = "&nbsp;"
		If le < l Then
			For i = 1 To l-le
				tmp = tmp & f
			Next
		End If
		Fill = s & tmp
	End Function

	' 返回正则验证结果
	' 调用: Test(内容, 验证规则)
	' 返回: 是否匹配 true/false
	Public Function Test(ByVal pStr, ByVal pRule)
		Dim lvRule
		Select Case Lcase(pRule)
			Case "date"		[Test] = isDate(pStr) : Exit Function
			Case "idcard"	[Test] = isIDCard(pStr) : Exit Function
			Case "number"	[Test] = isNumeric(pStr) : Exit Function
			Case "english"	lvRule = "^[A-Za-z]+$"
			Case "chinese"	lvRule = "^[\u4e00-\u9fa5]+$"
			Case "username"	lvRule = "^[a-zA-Z]\w{2,19}$"
			Case "email"	lvRule = "^\w+([-+\.]\w+)*@(([\da-zA-Z][\da-zA-Z-]{0,61})?[\da-zA-Z]\.)+([a-zA-Z]{2,4}(?:\.[a-zA-Z]{2})?)$"
			Case "int"		lvRule = "^[-\+]?\d+$"
			Case "double"	lvRule = "^[-\+]?\d+(\.\d+)?$"
			Case "price"	lvRule = "^\d+(\.\d+)?$"
			Case "zip"		lvRule = "^\d{6}$"
			Case "qq"		lvRule = "^[1-9]\d{4,9}$"
			Case "phone"	lvRule = "^((\(\+?\d{2,3}\))|(\+?\d{2,3}\-))?(\(0?\d{2,3}\)|0?\d{2,3}-)?[1-9]\d{4,7}(\-\d{1,4})?$"
			Case "mobile"	lvRule = "^(\+?\d{2,3})?0?1(3\d|47|5\d|8[056789])\d{8}$"
			Case "url"		lvRule = "^(?:(https|http|ftp|rtsp|mms)://(?:([\w!~\*'\(\).&=\+\$%-]+)(?::([\w!~\*'\(\).&=\+\$%-]+))?@)?)?((?:(?:(?:25[0-5]|2[0-4]\d|(?:1\d|[1-9])?\d)\.){3}(?:25[0-5]|2[0-4]\d|(?:1\d|[1-9])?\d))|(?:(?:(?:[\da-zA-Z][\da-zA-Z-]{0,61})?[\da-zA-Z]\.)+(?:[a-zA-Z]{2,4}(?:\.[a-zA-Z]{2})?)|localhost))(?::(\d{1,5}))?([#\?/].*)?$"
			Case "domain"	lvRule = "^(([\da-zA-Z][\da-zA-Z-]{0,61})?[\da-zA-Z]\.)+([a-zA-Z]{2,4}(?:\.[a-zA-Z]{2})?)$"
			Case "ip"		lvRule = "^((25[0-5]|2[0-4]\d|(1\d|[1-9])?\d)\.){3}(25[0-5]|2[0-4]\d|(1\d|[1-9])?\d)$"
			Case Else lvRule = pRule
		End Select
		Test = RegTest(CStr(pStr), lvRule)
	End Function

	' 依规则验证字符串是否存在匹配成功的结果
	' 调用: RegTest(内容, 搜寻规则)
	' 返回: true 匹配 false 不匹配
	Public Function RegTest(ByVal pStr, ByVal pRule)
		If IsNull(pStr) Then RegTest = False : Exit Function
		coRegExp.Pattern = pRule
		RegTest = coRegExp.Test(CStr(pStr))
		coRegExp.Pattern = ""
	End Function

	' 依规则取得结果集
	' 调用: RegMatch(内容, 搜寻规则)
	' 返回: 返回符事规则的结果集对像
	Public Function RegMatch(ByVal pStr, ByVal pRule)
		coRegExp.Pattern = pRule
		Set RegMatch = coRegExp.Execute(pStr)
		coRegExp.Pattern = ""
	End Function

	' 正则替换
	' 调用: RegReplace(内容, 规则, 结果)
	' 返回: 替换后的内容
	Public Function RegReplace(ByVal pStr, ByVal pRule, Byval pResult)
		RegReplace = RegReplace_(pStr, pRule, pResult, 0)
	End Function

	' 正则替换多行模式 (来自Easp2.2)
	' 调用: RegReplace(内容, 规则, 结果)
	' 返回: 替换后的内容
	Function RegReplaceM(ByVal pStr, ByVal pRule, Byval pResult)
		RegReplaceM = RegReplace_(pStr, pRule, pResult, 1)
	End Function

	' 正则表达式特殊字符转义 (来自Easp2.2)
	Function RegEncode(ByVal s)
		Dim re,i
		re = Split("\,$,(,),*,+,.,[,?,^,{,|",",")
		For i = 0 To Ubound(re)
			s = Replace(s,re(i),"\"&re(i))
		Next
		RegEncode = s
	End Function

	' 替换正则表达式编组 (来自Easp2.2)
	Function ReplacePart(ByVal txt, ByVal rule, ByVal part, ByVal replacement)
		If Not RegTest(txt, rule) Then
			ReplacePart = "[not match]"
			Exit Function
		End If
		Dim Match,i,j,ma,pos,uleft,ul
		i = Int(Mid(part,2))-1
		Set Match = RegMatch(txt,rule)(0)
		For j = 0 To Match.SubMatches.Count-1
			ma = Match.SubMatches(j)
			pos = Instr(txt,ma)
			If pos > 0 Then
				ul = Left(txt,pos-1)
				txt = Mid(txt,Len(ul)+1)
				If i = j Then
					ReplacePart = uleft & ul & Replace(txt,ma,replacement,pos-len(ul),1,0)
					Exit For
				End If
				uleft = uleft & ul & ma
				txt = Mid(txt, Len(ma)+1)
			End If
		Next
		Set Match = Nothing
	End Function


	' 特殊字符编码 (来自Easp2.2)
	' 调用: Escape(需要编码的字符)
	' 返回: Unicode编码后的字符
	' 说明: 把待编码的字符中的特殊字符按Unicode编码(对中文进行重新编码)
	Public Function Escape(ByVal pStr)
		If isNull(pStr) Then Escape = "" : Exit Function
		Dim i,c,a,s : s = ""
		For i = 1 To Len(pStr)
			c = Mid(pStr,i,1)
			a = ASCW(c)
			If (a>=48 and a<=57) or (a>=65 and a<=90) or (a>=97 and a<=122) Then
				s = s & c
			ElseIf InStr("@*_+-./",c)>0 Then
				s = s & c
			ElseIf a>0 and a<16 Then
				s = s & "%0" & Hex(a)
			ElseIf a>=16 and a<256 Then
				s = s & "%" & Hex(a)
			Else
				s = s & "%u" & Hex(a)
			End If
		Next
		Escape = s
	End Function

	' 特殊字符解码 (来自Easp2.2)
	' 调用: Ebody.UnEscape(需要解码的字符)
	' 返回: 对Unicode编码反编码后的字符
	' 说明: 把Unicode编码转成当前字符编码所对应的字符
	Public Function UnEscape(ByVal pStr)
		If isNull(pStr) Then UnEscape = "" : Exit Function
		Dim x, s
		x = InStr(pStr,"%")
		s = ""
		Do While x>0
			s = s & Mid(pStr,1,x-1)
			If LCase(Mid(pStr,x+1,1))="u" Then
				s = s & ChrW(CLng("&H"&Mid(pStr,x+2,4)))
				pStr = Mid(pStr,x+6)
			Else
				s = s & Chr(CLng("&H"&Mid(pStr,x+1,2)))
				pStr = Mid(pStr,x+3)
			End If
			x = InStr(pStr,"%")
		Loop
		UnEscape = s & pStr
	End Function

	' 处理字符串中的Javascript特殊字符 (来自Easp2.2)
	' 调用: Ebody.JsEncode(字符串)
	' 例子: Ebody.jsEncode("The Path is : '\test\path'") 返回: The Path is : \'\\test\\path\'
	Public Function JsEncode(ByVal s)
		If isNull(s) Then JsEncode = "" : Exit Function
		Dim arr1, arr2, i, j, c, p, t
		arr1 = Array(&h27,&h22,&h5C,&h2F,&h08,&h0C,&h0A,&h0D,&h09)
		arr2 = Array(&h27,&h22,&h5C,&h2F,&h62,&h66,&h6E,&h72,&h749)
		For i = 1 To Len(s)
			p = True
			c = Mid(s, i, 1)
			For j = 0 To Ubound(arr1)
				If c = Chr(arr1(j)) Then
					t = t & "\" & Chr(arr2(j))
					p = False
					Exit For
				End If
			Next
			If p Then 
				Dim a
				a = AscW(c)
				If a > 31 And a < 127 Then
					t = t & c
				ElseIf a > -1 Or a < 65535 Then
					t = t & "\u" & String(4 - Len(Hex(a)), "0") & Hex(a)
				End If 
			End If
		Next
		JsEncode = t
	End Function

	' 生成javascript代码 (来自Easp2.2)
	' 说明: 对JsEncode的产生的js安全字符进行反编译, 并生成可执行的js代码
	Public Function JsCode(ByVal s)
		JsCode = Str("<{1} type=""text/java{1}"">{2}{3}{4}{2}</{1}>{2}", Array("sc"&"ript",vbCrLf,vbTab,s))
	End Function

	' 取一个随机数 (来自Easp2.2)
	' 调用: Rand(随机数最小值, 随机数最大值)
	' 返回: 随机数值
	Public Function Rand(ByVal pMin, ByVal pMax)
		Randomize(Timer) : Rand = Int((pMax - pMin + 1) * Rnd + pMin)
	End Function

	' 格式化数字 (来自Easp2.2)
	' 调用: toNumber(待转换的值, 小数位数)
	' 返回: 格式化为带有,号分隔的美式数值
	' 样例: toNumber(12345.678, 2) 返回 12,345,68
	Public Function ToNumber(ByVal pNum, ByVal pDic)
		ToNumber = FormatNumber(pNum,pDic,-1)
	End Function

	' 将数字转换为货币格式 (来自Easp2.2)
	' 调用: ToPrice(待转换的值)
	' 返回: 参数number中的值转换为带有货币符号的数值,何种货币符号取决于服务器语言区域相关选项的设置
	Public Function ToPrice(ByVal pNum)
		ToPrice = FormatCurrency(pNum,2,-1,0,-1)
	End Function

	' 将数字转换为百分比格式 (来自Easp2.2)
	' 调用: ToPercent(待转换的值)
	' 返回: 参数number中的值转换为带有两位小数的百分比样式
	Public Function ToPercent(ByVal pNum)
		ToPercent = FormatPercent(pNum,2,-1)
	End Function

	' 取字符隔开的左段 (来自Easp2.2)
	' 调用: Ebody.CLeft("abc.efg", ".")
	' 返回: abc
	Public Function CLeft(ByVal pStr, ByVal pM)
		CLeft = GetStr_(pStr,pM,0)
	End Function

	' 取字符隔开的右段 (来自Easp2.2)
	' 调用: Ebody.CLeft("abc.efg", ".")
	' 返回: efg
	Public Function CRight(ByVal pStr, ByVal pM)
		CRight = GetStr_(pStr,pM,1)
	End Function

	' 取得文件或文件夹在服务器上的物理存放位置(支持通配符*和?)
	' 调用: MapPath(虚拟文件路径)
	' 返回: 物理地址
	Public Function MapPath(pPath)
		MapPath = AbsPath_(pPath)
	End Function

	'-------------------------
	' 验证类函数(Is打头)
	'-------------------------

	' 检测文件或文件夹或驱动器(磁盘)是否存在
	' 调用: IsExists(对像路径[支持相对与绝对])
	' 返回: true 存在, false 不存在
	' 样例: IsExists("/dev/demo/abc.asp")
	' 说明: 地址第一个符号是/, 则位置从根目录开始(即:绝对地址), 不带则位置从当前文件所在文件夹开始(即:相对地址)
	Public Function IsExists(ByVal pPath)		
		If IsFile(pPath) Or IsFolder(pPath) Or IsDrive(pPath) Then IsExists = True Else IsExists = False
	End Function

	' 检测文件是否存在
	' 调用: IsFile(文件路径[支持相对与绝对])
	' 返回: true 存在, false 不存在
	' 样例: IsFile("/dev/demo/abc.asp")
	' 说明: 地址第一个符号是/, 则位置从根目录开始(即:绝对地址), 不带则位置从当前文件所在文件夹开始(即:相对地址)
	Public Function IsFile(ByVal pFilePath)
		pFilePath = AbsPath_(pFilePath)
		If coFso.FileExists(pFilePath) Then IsFile = True Else IsFile = False
	End Function
	
	' 检测文件夹是否存在
	' 调用: IsFolder(目录路径[支持相对与绝对])
	' 返回: true 存在, false 不存在
	' 样例: IsFolder("/common/system")
	' 说明: 地址第一个符号是/, 则位置从根目录开始(即:绝对地址), 不带则位置从当前文件所在文件夹开始(即:相对地址)
	Public Function IsFolder(ByVal pFolderPath)
		pFolderPath = AbsPath_(pFolderPath)
		If coFso.FolderExists(pFolderPath) Then IsFolder = True Else IsFolder = False
	End Function

	' 检测驱动器是否存在
	' 调用: IsDrive(盘符)
	' 返回: true 存在, false 不存在
	' 样例: IsDrive("d:")
	Public Function IsDrive(ByVal pDrive)
		pDrive = AbsPath_(pDrive)
		If coFSO.DriveExists(pDrive) Then IsDrive = True Else IsDrive = False
	End Function

	' 检测服务器组件是否安装 (来自Easp2.2)
	' 调用: IsInstall(组件名称)
	' 返回: true/false
	Public Function IsInstall(Byval pStr)
		'On Error Resume Next : Err.Clear()
		IsInstall = False
		Dim loObj : Set obj = Server.CreateObject(pStr)
		If Err.Number = 0 Then IsInstall = True
		Set loObj = Nothing : Err.Clear()
	End Function

	' 检测Ebody插件是否载入
	' 调用: IsLoad(插件名称)
	' 返回: true/false
	Public Function IsLoad(Byval pClassName)
		'On Error Resume Next : Err.Clear()		
		IsLoad = False
		' 判断是否为扩展子类
		Dim lvBaseClass : lvBaseClass = pClassName
		If InStr(lvBaseClass, ".") Then lvBaseClass = CLeft(lvBaseClass, ".")
		' 验证返回
		If Eval("IsObject(" & lvBaseClass & ")") Then
			If Not (InStr("nothing,empty", Eval("LCase(TypeName(" & pClassName & "))"))>0) And Len(pClassName)>0 Then IsLoad = True
		End If
	End Function

	' 验证并记录系统错误
	Public Function IsErr()
		If Err.Number<>0 Then
			IsErr = True
			cvErrCode = Err.Number
			cvErrMsg = cvErrMsg & ":" & "(" & Err.Description & ")"
		Else
			IsErr = False
		End If
		Err.Clear()
	End Function

	'-------------------------
	' 系统取值类函数(Get打头)
	'-------------------------

	' 依相对路径取得文件绝对物理路径
	' 调用: GetPathAbs(文件路径[支持相对与绝对][以/开头代表从根目录开始,否则从当前页面目录开始定位])
	' 返回: 文件基于根目录的绝对物理路径
	' 例子: 如当前网站根目录是: d:\tony\ebody110
	'		调用: GetPathAbs("\Server\Plugin\ebody.music.asp")
	'		返回: d:\tony\ebody110\Server\Plugin\ebody.music.asp
	' 说明: 如果你并不确定你的上层目录是什么,但只想从根目录开始定位,此时GetPathAbs就有用武之地了
	'		类似于Server.MapPath()函数,此为其加强版
	Public Function GetPathAbs(ByVal pPath)
		If (Mid(pPath, 2, 1)<>":") And (InStr(pPath, "://")<1 Or InStr(pPath, Request.ServerVariables("Server_Name"))<1) Then
			'GetPathAbs = VRoot & Replace(pPath, "/", "\")
			GetPathAbs = Server.MapPath(pPath)
		Else
			GetPathAbs = pPath
		End If
	End Function

	' 依相对路径取得页面绝对网络路径
	' 调用: GetUrlAbs(文件路径[以/开头代表从根网址开始,否则从当前页面目录开始定位])
	' 返回: 文件的绝对网络路径
	' 例子: 如当前网站主页是: http://localhost
	'		调用: GetUrlAbs("/Server/Plugin/demo.asp")
	'		返回: http://localhost/Server/Plugin/demo.asp
	' 说明: 如果你并不确定你的上层目录是什么,但只想从根目录开始定位,此时GetUrlAbs就有用武之地了
	Public Function GetUrlAbs(ByVal pPath)
		pPath = Replace(pPath,"\","/")
		If (Mid(pPath, 2, 1)<>":") And (InStr(pPath, "://")<1 Or InStr(pPath, Request.ServerVariables("Server_Name"))<1) Then
			If InStr(pPath,"/")=1 Then
				GetUrlAbs = VHome & Replace(pPath, "\", "/")
			Else				
				Dim lvPath : lvPath = Request.ServerVariables("Script_Name")
				GetUrlAbs = Home & Left(lvPath, InStrRev(lvPath, "/") - 1) & "/" & pPath
			End If
		Else
			GetUrlAbs = pPath
		End If
	End Function	

	' 脚本执行时间(毫秒)
	' 调用: GetScriptTime(Timer())
	' 返回: 脚本执行时间(保留3位小数)
	Public Function GetScriptTime(pTime)
		If pTime="" Or pTime="0" Or pTime=Empty Then pTime = cvSysTime
		GetScriptTime = FormatNumber((Timer() - pTime) * 1000, 3, -1)
	End Function

	' 获取用户IP地址
	' 调用: GetClientIP()
	' 返回: 192.168.0.1
	' 说明: 取得请求者终端来源IP地址(取终端来源IP,如有代理服务器,则取代理的客户端IP)
	Public Function GetClientIP()
		Dim lvIP, x, y
		x = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
		y = Request.ServerVariables("REMOTE_ADDR")
		lvIP = IIF(IsNull(x) or lCase(x)="unknown",y,x)
		If InStr(lvIP,".")=0 Then lvIP = "0.0.0.0"
		GetClientIP = lvIP
	End Function

	' 查询CodePage和Charset对照值(easp 2.2)
	' 调用: GetCharCode(代码或字符集名)
	' 样例: GetCharCode("utf-8")
	' 返回: 65001
	Public Function GetCharCode(ByVal pS)
		Dim co, ch, cn, cnf, i
		co = Array(708,720,28596,1256,1257,852,28592,1250,936,936,950,862,66,874,932,949,1251,1252,1253,1254,1255,1258,20866,21866,28595,28597,28598,38598,50932,51932,52936,65001)
		ch = Array("ASMO-708","DOS-720","iso-8859-6","windows-1256","windows-1257","ibm852","iso-8859-2","windows-1250","gb2312","gbk","big5","DOS-862","cp866","windows-874","shift_jis","ks_c_5601-1987","windows-1251","iso-8859-1","windows-1253","iso-8859-9","windows-1255","windows-1258","koi8-r","koi8-ru","iso-8859-5","iso-8859-7","iso-8859-8","iso-8859-8-i","_autodetect","euc-jp","hz-gb-2312","utf-8")
		cn = Array("阿拉伯字符 (ASMO 708)","阿拉伯字符 (DOS)","阿拉伯字符 (ISO)","阿拉伯字符 (Windows)","波罗的海字符 (Windows)","中欧字符 (DOS)","中欧字符 (ISO)","中欧字符 (Windows)","简体中文 (GB2312)","简体中文 (GBK)","繁体中文 (Big5)","希伯来字符 (DOS)8","西里尔字符 (DOS)","泰语 (Windows)","日语 (Shift-JIS)","朝鲜语","西里尔字符 (Windows)","西欧字符","希腊字符 (Windows)","土耳其字符 (Windows)","希伯来字符 (Windows)","越南字符 (Windows)","西里尔字符 (KOI8-R)","西里尔字符 (KOI8-U)","西里尔字符 (ISO)","希腊字符 (ISO)","希伯来字符 (ISO-Visual)","希伯来字符 (ISO-Logical)","日语 (自动选择)","日语 (EUC)","简体中文 (HZ)","Unicode (UTF-8)")
		If Instr(pS,":") Then
			s = CLeft(pS,":")
			cnf = True
		End If
		If IsNumeric(pS) Then
			For i = 0 To Ubound(co)
				If co(i) = s Then
					GetCharCode = IIF(cnf,cn(i),ch(i))
					Exit Function
				End If
			Next
			GetCharCode = "utf-8"
		Else
			For i = 0 To Ubound(ch)
				If ch(i) = LCase(pS) Then
					GetCharCode = IIF(cnf,cn(i),co(i))
					Exit Function
				End If
			Next
			GetCharCode = 65001
		End If
	End Function

	' 取得值类型
	' 调用: GetVarType(值)
	' 返回: 值的类型
	Public Function GetVarType(ByVal pValue)
		Select Case VarType(pValue)
			Case vbArray
				GetVarType = "Array"
			Case vbBoolean
				GetVarType = "Boolean"
			Case vbByte
				GetVarType = "Byte"
			Case vbCurrency
				GetVarType = "Currency"
			Case vbDataObject
				GetVarType = "DataObject"
			Case vbDate
				GetVarType = "Date"
			Case vbDecimal
				GetVarType = "Decimal"
			Case vbDouble
				GetVarType = "Double"
			Case vbEmpty
				GetVarType = "Empty"
			Case vbError
				GetVarType = "Error"
			Case vbInteger
				GetVarType = "Integer"
			Case vbLong
				GetVarType = "Long"
			Case vbNull
				GetVarType = "Null"
			Case vbObject
				GetVarType = "Object"
			Case vbSingle
				GetVarType = "Single"
			Case vbString
				GetVarType = "String"
			Case vbUserDefinedType
				GetVarType = "UserDefinedType"
			Case vbVariant
				GetVarType = "Variant"
			Case Else
				GetVarType = "Unknow"
		End Select
	End Function	

	' 获取一个Cookies值 (来自Easp2.2)
	' 调用: GetCookie(Cookie名称[:关键字])
	' 返回: Cookies值字符串
	Public Function GetCookie(ByVal pStr)
		Dim p,t,lvCookie
		If Instr(pStr,">") > 0 Then
			p = CLeft(pStr,">")
			s = CRight(pStr,">")
		End If
		If Instr(pStr,":")>0 Then
			t = CRight(pStr,":")
			s = CLeft(pStr,":")
		End If
		If Has(p) And Has(pStr) Then
			If Response.Cookies(p).HasKeys Then
				lvCookie = Request.Cookies(p)(pStr)
			End If
		ElseIf Has(pStr) Then
			lvCookie = Request.Cookies(pStr)
		Else
			GetCookie = "" : Exit Function
		End If
		If IsNull(lvCookie) Then GetCookie = "": Exit Function
		If  cvCookieEncode Then
			If LCase(InStr(GetFolderFiles_(cvCoreClassPath)), "aes")>0 Then
				Use("Aes") : lvCookie = Aes.Decode(lvCookie)	' ***使用到加密, AES类待完善***
			End If
		End If
		'GetCookie = Safe(lvCookie,t)
		GetCookie = lvCookie
	End Function

	' 获取一个缓存记录 (来自Easp2.2)
	' 调用: GetApp(缓存名称)
	' 返回: 系统缓存值
	Public Function GetApp(ByVal pAppName)
		If IsNull(pAppName) Then GetApp = Empty : Exit Function
		If IsObject(Application(pAppName)) Then
			Set GetApp = Application(pAppName)
		Else
			GetApp = Application(pAppName)
		End If
	End Function	

	' 将HTML代码转换为文本实体 (来自Easp2.2)
	' 说明: 将加码参数中的字符串, 以方便阅读的HTML格式显示出来. 例如: 在页面上显示出包含HTML代码的原内容
	Public Function HtmlEncode(ByVal s)
		If Has(s) Then
			s = Replace(s, Chr(38), "&#38;")
			s = Replace(s, "<", "&lt;")
			s = Replace(s, ">", "&gt;")
			s = Replace(s, Chr(39), "&#39;")
			s = Replace(s, Chr(32), "&nbsp;")
			s = Replace(s, Chr(34), "&quot;")
			s = Replace(s, Chr(9), "&nbsp;&nbsp; &nbsp;")
			s = Replace(s, vbCrLf, "<br />")
		End If
		HtmlEncode = s
	End Function

	' 将HTML文本转换为HTML代码 (来自Easp2.2)
	' 说明: 将解码其中的特殊字符, 与HtmlEncode对应
	Public Function HtmlDecode(ByVal s)
		If Has(s) Then
			s = regReplace(s, "<br\s*/?\s*>", vbCrLf)
			s = Replace(s, "&nbsp;&nbsp; &nbsp;", Chr(9))
			s = Replace(s, "&quot;", Chr(34))
			s = Replace(s, "&nbsp;", Chr(32))
			s = Replace(s, "&#39;", Chr(39))
			s = Replace(s, "&apos;", Chr(39))
			s = Replace(s, "&gt;", ">")
			s = Replace(s, "&lt;", "<")
			s = Replace(s, "&amp;", Chr(38))
			s = Replace(s, "&#38;", Chr(38))
		End If
		HtmlDecode = s
	End Function

	' 过滤HTML标签 (来自Easp2.2)
	' 说明: 将过滤参数字符串中的所有HTML标签, 返回纯文本
	Public Function HtmlFilter(ByVal s)
		If Has(s) Then
			If IsNull(s) Then HtmlFilter = "" : Exit Function
			s = regReplace(s,"<[^>]+>","")
			s = Replace(s, ">", "&gt;")
			s = Replace(s, "<", "&lt;")
		End If
		HtmlFilter = s
	End Function

	' 内容编码转换
	' 调用: CharsetTo(要转换的内容, 要转换成的目标字符集)
	Public Function CharsetTo(ByVal pStr, ByVal pCharset) 
		CharsetTo = CharsetTo_(pStr, pCharset)
	End Function

	' 格式化日期时间 (来自Easp2.2)
	Public Function DateTime(ByVal iTime, ByVal iFormat)
		If IsNull(iTime) Then DateTime = "" : Exit Function
		If Not IsDate(iTime) Then DateTime = "Date Error" : Exit Function
		If Instr(",0,1,2,3,4,",","&iFormat&",")>0 Then DateTime = FormatDateTime(iTime,iFormat) : Exit Function
		Dim diffs,diffd,diffw,diffm,diffy,dire,before,pastTime
		Dim iYear, iMonth, iDay, iHour, iMinute, iSecond,iWeek,tWeek
		Dim iiYear, iiMonth, iiDay, iiHour, iiMinute, iiSecond,iiWeek
		Dim iiiWeek, iiiMonth, iiiiMonth
		Dim SpecialText, SpecialTextRe,i,t
		iYear = right(Year(iTime),2) : iMonth = Month(iTime) : iDay = Day(iTime)
		iHour = Hour(iTime) : iMinute = Minute(iTime) : iSecond = Second(iTime)
		iiYear = Year(iTime) : iiMonth = right("0"&Month(iTime),2)
		iiDay = right("0"&Day(iTime),2) : iiHour = right("0"&Hour(iTime),2)
		iiMinute = right("0"&Minute(iTime),2) : iiSecond = right("0"&Second(iTime),2)
		tWeek = Weekday(iTime)-1 : iWeek = Array("日","一","二","三","四","五","六")
		If isDate(iFormat) or isNull(iFormat) Then
			If isNull(iFormat) Then : iFormat = Now() : pastTime = true : End If
			dire = "后" : If DateDiff("s",iFormat,iTime)<0 Then : dire = "前" : before = True : End If
			diffs = Abs(DateDiff("s",iFormat,iTime))
			diffd = Abs(DateDiff("d",iFormat,iTime))
			diffw = Abs(DateDiff("ww",iFormat,iTime))
			diffm = Abs(DateDiff("m",iFormat,iTime))
			diffy = Abs(DateDiff("yyyy",iFormat,iTime))
			If diffs < 60 Then DateTime = "刚刚" : Exit Function
			If diffs < 1800 Then DateTime = Int(diffs\60) & "分钟" & dire : Exit Function
			If diffs < 2400 Then DateTime = "半小时"  & dire : Exit Function
			If diffs < 3600 Then DateTime = Int(diffs\60) & "分钟" & dire : Exit Function
			If diffs < 259200 Then
				If diffd = 3 Then DateTime = "3天" & dire & " " & iiHour & ":" & iiMinute : Exit Function
				If diffd = 2 Then DateTime = IIF(before,"前天 ","后天 ") & iiHour & ":" & iiMinute : Exit Function
				If diffd = 1 Then DateTime = IIF(before,"昨天 ","明天 ") & iiHour & ":" & iiMinute : Exit Function
				DateTime = Int(diffs\3600) & "小时" & dire : Exit Function
			End If
			If diffd < 7 Then DateTime = diffd & "天" & dire & " " & iiHour & ":" & iiMinute : Exit Function
			If diffd < 14 Then
				If diffw = 1 Then DateTime = IIF(before,"上星期","下星期") & iWeek(tWeek) & " " & iiHour & ":" & iiMinute : Exit Function
				If Not pastTime Then DateTime = diffd & "天" & dire : Exit Function
			End If
			If Not pastTime Then
				If diffd < 31 Then
					If diffm = 2 Then DateTime = "2个月" & dire : Exit Function
					If diffm = 1 Then DateTime = IIF(before,"上个月","下个月") & iDay & "日" : Exit Function
					DateTime = diffw & "星期" & dire : Exit Function
				End If
				If diffm < 36 Then
					If diffy = 3 Then DateTime = "3年" & dire : Exit Function
					If diffy = 2 Then DateTime = IIF(before,"前年","后年") & iMonth & "月" : Exit Function
					If diffy = 1 Then DateTime = IIF(before,"去年","明年") & iMonth & "月" : Exit Function
					DateTime = diffm & "个月" & dire : Exit Function
				End If
				DateTime = diffy & "年" & dire : Exit Function
			Else
				iFormat = "yyyy-mm-dd hh:ii"
			End If
		End If
		iiWeek = Array("Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday")
		iiiWeek = Array("Sun","Mon","Tue","Wed","Thu","Fri","Sat")
		iiiMonth = Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
		iiiiMonth = Array("January","February","March","April","May","June","July","August","September","October","November","December")
		SpecialText = Array("y","m","d","h","i","s","w")
		SpecialTextRe = Array(Chr(0),Chr(1),Chr(2),Chr(3),Chr(4),Chr(5),Chr(6))
		For i = 0 To 6 : iFormat = Replace(iFormat,"\"&SpecialText(i), SpecialTextRe(i)) : Next
		t = Replace(iFormat,"yyyy", iiYear) : t = Replace(t, "yyy", iiYear)
		t = Replace(t, "yy", iYear) : t = Replace(t, "y", iiYear)
		t = Replace(t, "mmmm", Replace(iiiiMonth(iMonth-1),"m",Chr(1))) : t = Replace(t, "mmm", iiiMonth(iMonth-1))
		t = Replace(t, "mm", iiMonth) : t = Replace(t, "m", iMonth)
		t = Replace(t, "dd", iiDay) : t = Replace(t, "d", iDay)
		t = Replace(t, "hh", iiHour) : t = Replace(t, "h", iHour)
		t = Replace(t, "ii", iiMinute) : t = Replace(t, "i", iMinute)
		t = Replace(t, "ss", iiSecond) : t = Replace(t, "s", iSecond)
		t = Replace(t, "www", iiiWeek(tWeek)) : t = Replace(t, "ww", iiWeek(tWeek))
		t = Replace(t, "w", iWeek(tWeek))
		For i = 0 To 6 : t = Replace(t, SpecialTextRe(i),SpecialText(i)) : Next
		DateTime = t
	End Function

	' 取指定长度的随机字符串 (来自Easp2.2)
	Public Function RandStr(ByVal cfg)
		Dim a, p, l, t, reg, m, mi, ma
		cfg = Replace(Replace(Replace(cfg,"\<",Chr(0)),"\>",Chr(1)),"\:",Chr(2))
		a = ""
		If RegTest(cfg, "(<\d+>|<\d+-\d+>)") Then
			t = cfg
			p = MParam(cfg)
			If Not isNull(p(1)) Then
				a = p(1) : t = p(0) : p = ""
			End If
			Set reg = RegMatch(cfg, "(<\d+>|<\d+-\d+>)")
			For Each m In reg
				p = m.SubMatches(0)
				l = Mid(p,2,Len(p)-2)
				If RegTest(l,"^\d+$") Then
					t = Replace(t,p,RandStr_(l,a),1,1)
				Else
					mi = CLeft(l,"-")
					ma = CRight(l,"-")
					t =  Replace(t,p,Rand(mi, ma),1,1)
				End If
			Next
			Set reg = Nothing
		ElseIf RegTest(cfg,"^\d+-\d+$") Then
			mi = CLeft(cfg,"-")
			ma = CRight(cfg,"-")
			t = Rand(mi, ma)
		ElseIf RegTest(cfg, "^(\d+)|(\d+:.)$") Then
			l = cfg : p = MParam(cfg)
			If Not isNull(p(1)) Then
				a = p(1) : l = p(0) : p = ""
			End If
			t = RandStr_(l, a)
		Else
			t = cfg
		End If
		RandStr = Replace(Replace(Replace(t,Chr(0),"<"),Chr(1),">"),Chr(2),":")
	End Function

'--------------------------------------------------------------------------------
'-- Sub
'--------------------------------------------------------------------------------

	'-------------------------
	' 基础功能类过程
	'-------------------------

	' 关闭指定对象
	' 调用: Close(待关闭的对象引用或对象名字串)
	Public Sub Close(ByRef pObj)
		'On Error Resume Next
		If IsObject(pObj) Then
			Select Case TypeName(pObj)
			Case "Connection"
				If pObj.State = 1 Then pObj.Close()
				Set pObj = Nothing
			Case "Recordset"
				If pObj.State = 1 Then pObj.Close()
				Set pObj = Nothing
			Case Else
				Set pObj = Nothing
			End Select
		Else
			' 动态关闭对像
			Execute("Set " & pObj & " = Nothing")
		End If
	End Sub

	' 关闭所有核心实例对象
	' 说明: 需要fso支持
	Public Sub CloseCores()
		' 得到核心类库目录文件名清单
		Dim lvCoreFileList : lvCoreFileList = GetFolderFiles_(cvCoreClassPath)
		' 将文件清单转换为类库变量名列表, 格式如: db,tpl,http
		Dim lvCoreClassList : lvCoreClassList = GetClassList_(lvCoreFileList)
		' 逐各释放所有核心类内存
		CloseObjFromList_(lvCoreClassList)
	End Sub

	' 关闭所有扩展实例对象
	' 说明: 需要fso支持
	Public Sub CloseExts()
		' --- 以下代码取消也OK,待参考一段时间 ---
		' 因为所有扩展子类对像都是存在ext对像中的,所以仅需释放ext对像即全部释放
		' -----------------------
		' 得到扩展类库目录文件名清单
		Dim lvExtFileList : lvExtFileList = GetFolderFiles_(cvExtClassPath)
		' 将文件清单转换为类库变量名列表,格式如: ext.music,ext.demo,ext.play
		Dim lvExtClassList : lvExtClassList = GetClassList_(lvExtFileList)
		' 将类名转换为扩展类下的引用类名
		Dim lvClass, lvClasses, lvExtClassListTmp
		lvClasses = Split(lvExtClassList, ",")
		For Each lvClass In lvClasses
			If Not Len(lvExtClassListTmp) Then
				lvExtClassListTmp = "ext." & lvClass
			Else
				lvExtClassListTmp = lvExtClassListTmp & "," & "ext." & lvClass
			End If
		Next
		lvExtClassList = lvExtClassListTmp
		' 逐各释放所有扩展子类对像内存
		If IsObject(ext) Then CloseObjFromList_(lvExtClassList)
		' ------------------------

		' 释放扩展基类ext对像的内存,其内的所有扩展子类实例对像同时也会释放
		' 但这样做的结果是不能用isobject去判断其对像,会影响到isload的正常使用
		'Set ext = Nothing
	End Sub

	' 关闭所有Ebody实例对象,释放内存
	' 调用: CloseAll
	Public Sub CloseAll()
		' 1 先释放扩展子类
		Call CloseExts
		' 2 现释放核心基类
		Call CloseCores
	End Sub

	' 将文件动态载入到内存中
	' 调用: Include(文件路径[支持相对与绝对])
	' 说明: 此方法支持带有<inclued>标签的文件的无限级载入
	'		载入内存后, 即可调用其功能, 相当于扩展接口
	Public Sub Include(ByVal pFilePath)
		Dim lvContent
		pFilePath = AbsPath_(pFilePath)
		lvContent = GetFileAll_(pFilePath)
		lvContent = Replace(lvContent, "<"&"%", "")
		lvContent = Replace(lvContent, "%"&">", "")
		' 动态执行脚本, 并载入到内存
		ExecuteGlobal(lvContent)
		' 记录错误
		ErrLog_("Include")
	End Sub

	' 不缓存页面信息
	' 调用: Ebody.NoCache
	Public Sub NoCache()
		Response.Buffer = True
		Response.Expires = 0
		Response.ExpiresAbsolute = Now() - 1
		Response.CacheControl = "no-cache"
		Response.AddHeader "Expires",Date()
		Response.AddHeader "Pragma","no-cache"
		Response.AddHeader "Cache-Control","private, no-cache, must-revalidate"
	End Sub

	' 加载核心类
	' 调用: Use(基类的类型名)
	' 样例: Ebody.Use "Tpl"
	' 应用:	Ebody.Tpl.Show
	' 说明: 加载的类文件必须是以ebody.开头的asp文件,加载的类名必须以ebody_开头
	Public Sub Use(ByVal pClassType)
		Dim lvClassFile, lvClassName
		' 定义核心类的类名
		lvClassName = "ebody_" & pClassType
		' 验证基类是否已经实例化
		If Eval("LCase(TypeName(" & pClassType & "))") = lvClassName Then Exit Sub
		' 验证基类是否已载入,对像已载入,则重新创建对像实例
		If Eval("IsObject(" & pClassType & ")") Then Execute("Set " & pClassType & " = New " & lvClassName) : Exit Sub
		' 依基类类型名,组构基类文件所在的路径
		lvClassFile = cvCoreClassPath & "\" & "ebody." & Lcase(pClassType) & ".asp"
		' 验证基类文件是否存在
		If Not IsFile(lvClassFile) Then Exit Sub
		' 载入基类文件到内存
		Call Include(lvClassFile)
		' 实例化基类
		Execute("Set " & pClassType & " = New " & lvClassName)
	End Sub

	' 加载扩展类
	' 调用: Extend(扩展子类的类型名)
	' 样例: ebody.Extend "Test"
	' 应用: ebody.ext.test.load "demo.html"
	'		ebody为核心类实例,ext为扩展类,test为ebody.ext下动态创建的扩展子类实例,load为扩展子类test中的方法
	' 说明: 可应用于插件功能
	'		扩展子类文件必须存放在系统指定的目录中
	'		扩展子类的命名必须符合系统规则,类文件名:ebody. + 子类名 + .文件扩展名,类名:ebody_ + 子类名
	'		系统从指定的扩展子类目录中寻找所有符合标准的子类文件,并从中提取其类名,将类名动态加入到扩展类ext中,
	'		用于存储动态生成的扩展类实例,这样就可方便前台直接引用此变量进行操作类
	'		------------------------------------------------------------------
	'		扩展类运行流程及说明:
	'		1 取得类库目录下所有类库文件名
	'		2 将文件名转换成类库变量名(此变量用于存入实例后的加载类实例)
	'		3 组构扩展类,依据上步中取得的类库变量,动态新增至扩展类的内容中(动态组构类内容)
	'		4 将动态组构好的扩展类代码载入内存(使用ExecuteGlobal)
	'		5 实例化扩展类,并保存到核心基类的变量EXT中(如:set ext=new ebody_ext),接下来在前台就可直接引用此变量来操作类
	'		6 动态载入扩展子类(如载入的是:ebody_test)
	'		7 动态实例化扩展子类(此时就可使用Extend方法进行加载了,如:Ebody.Extend "test" )
	'		8 执行完成以上7步,接下来就可以使用扩展类中的方法了(如:ebody.ext.test.Show)
	'		9 释放扩展类实例(如:ebody.ext.close "test")
	'------------------------------------------------------------------
	Public Sub Extend(ByVal pClassType)
		Dim lvClassFile, lvClassName
		' 定义标准的扩展子类类名
		lvClassName = "ebody_" & pClassType
		' 依扩展子类类型名,组构扩展子类文件路径
		lvClassFile = cvExtClassPath & "\" & "ebody." & LCase(pClassType) & ".asp"
		' 验证扩展子类文件是否存在
		If Not IsFile(lvClassFile) Then Exit Sub
		' 实例化扩展子类的基类ext(此时将扩展类目录中的所有扩展类变量动态加入扩展基类中,这里是关键点)
		If Not IsObject(ext) Then Call RegisterExtClass_
		' 验证扩展基类是否已经实例化
		If Not LCase(TypeName(ext))="ebody_ext" Then Exit Sub
		' 验证扩展子类是否已经实例化
		If Eval("LCase(TypeName(ext." & pClassType & "))")=LCase(lvClassName) Then Exit Sub	
		' ---- 以上验证都通过,说明扩展子类还没有实例化,则下面进行实例化 ---
		' 载入扩展子类文件到内存
		Call Include(lvClassFile)
		' 实例化扩展子类(子类引用是在ext基类下)
		' 说明: pClassType与扩展类中的变量要对应.如pClassType的值为tpl,那么在ext类中也应有tpl这个变量
		'		这里是由系统自动从扩展类目录中通过RegisterExtClass_分析取得的,
		Execute("Set ext." & pClassType & " = New " & lvClassName)
	End Sub

	'-------------------------
	' 设置对像类过程(Set/Remove打头)
	'-------------------------

	' 设置一个Cookies值 (来自Easp2.2)
	' 调用: SetCookie(Cookies的名称[:Cookies关键字的名称], Cookies的值, Cookies的选项) 
	' Cookies的选项: 
	' Cookies的存活期限、路径、域及安全设置，用数组参数可设置多个选项。可以是以下值：
	' 数值或时间值 - 设置Cookies的存活期限，如果是时间值则表示此Cookies到指定的时间失效，
	' 如果是数字则表示此Cookies可以存活的时间长度(以分钟为单位)，如果为数字0则表示此Cookies到浏览器关闭时失效。
	' 域名字符串 - 设置Cookies的有效域名。
	' 相对路径字符串 - 设置Cookies的有效路径。
	' 布尔值 - 设置Cookies的安全性。
	' 数组 - 可以同时设置以上4种Cookies选项。
	' 样例: call SetCookie("MyServerIP", "192.168.0.1", 30) '设置一个30分钟后失效的Cookies
	Public Sub SetCookie(ByVal pCookName, ByVal pCookValue, ByVal pCookCfg)
		Dim n,i,lvExp,lvDomain,lvPath,lvSecure
		If isArray(pCookCfg) Then
			For i = 0 To Ubound(pCookCfg)
				If isDate(pCookCfg(i)) Then
					lvExp = cDate(pCookCfg(i))
				ElseIf Test(pCookCfg(i),"int") Then
					If pCookCfg(i)<>0 Then lvExp = Now()+Int(pCookCfg(i))/60/24
				ElseIf Test(pCookCfg(i),"domain") or Test(pCookCfg(i),"ip") Then
					lvDomain = pCookCfg(i)
				ElseIf Instr(pCookCfg(i),"/")>0 Then
					lvPath = pCookCfg(i)
				ElseIf pCookCfg(i)="True" or pCookCfg(i)="False" Then
					lvSecure = pCookCfg(i)
				End If
			Next
		Else
			If isDate(pCookCfg) Then
				lvExp = cDate(pCookCfg)
			ElseIf Test(pCookCfg,"int") Then
				If pCookCfg<>0 Then lvExp = Now()+Int(pCookCfg)/60/24
			ElseIf Test(pCookCfg,"domain") or Test(pCookCfg,"ip") Then
				lvDomain = pCookCfg
			ElseIf Instr(pCookCfg,"/")>0 Then
				lvPath = pCookCfg
			ElseIf pCookCfg = "True" or pCookCfg = "False" Then
				lvSecure = pCookCfg
			End If
		End If
		If Has(pCookValue) Then
			If cvCookieEncode Then
				If LCase(InStr(GetFolderFiles_(cvCoreClassPath)), "aes")>0 Then
					Use("Aes") : pCookValue = Aes.Encode(pCookValue)	' ***使用到加密, AES类待完善***
				End If
			End If
		End If
		If Instr(pCookName,">")>0 Then
			n = CRight(pCookName,">")
			pCookName = CLeft(pCookName,">")
			Response.Cookies(pCookName)(n) = pCookValue
		Else
			Response.Cookies(pCookName) = pCookValue
		End If
		If Has(lvExp) Then Response.Cookies(pCookName).Expires = lvExp
		If Has(lvDomain) Then Response.Cookies(pCookName).Domain = lvDomain
		If Has(lvPath) Then Response.Cookies(pCookName).Path = lvPath
		If Has(lvSecure) Then Response.Cookies(pCookName).Secure = lvSecure
	End Sub

	' 删除一个Cookies值 (来自Easp2.2)
	' 调用: RemoveCookie(Cookies名称)
	Public Sub RemoveCookie(ByVal s)
		Dim p,t
		If Instr(s,">") > 0 Then
			p = CLeft(s,">")
			s = CRight(s,">")
		End If
		If Has(p) And Has(s) Then
			If Response.Cookies(p).HasKeys Then
				Response.Cookies(p)(s) = Empty
			End If
		ElseIf Has(s) Then
			Response.Cookies(s) = Empty
			Response.Cookies(s).Expires = Now()
		End If
	End Sub

	' 设置缓存记录 (来自Easp2.2)
	' 调用: SetApp(缓存名称, 缓存数据)
	' 说明: 调用此方法将设置一个缓存记录的值(缓存(Application)是一种所有用户均可访问的全局变量，有点类似于所有用户共同的Session)
	Public Sub SetApp(ByVal pAppName, ByRef pAppData)
		Application.Lock
		If IsObject(pAppData) Then
			Set Application(pAppName) = pAppData
		Else
			Application(pAppName) = pAppData
		End If
		Application.UnLock
	End Sub

	' 删除一个缓存记录 (来自Easp2.2)
	' 调用: RemoveApp(缓存名称)
	Public Sub RemoveApp(ByVal pAppName)
		Application.Lock
		Application(pAppName) = Empty
		Application.UnLock
	End Sub

	' 记录系统错误
	Public Sub Log(ByVal pMsg)
		ErrLog_(pMsg)
	End Sub	

	' 在服务器端输出javascript执行代码到客户端 (来自Easp2.2)
	' 说明: 此功能可动态的在服务器端自定义生成js代码, 传给客户端执行
	Public Sub Js(ByVal s)
		Response.write JsCode(s)
	End Sub

	' 服务器端输出javascript弹出消息 (来自Easp2.2)
	Public Sub Alert(ByVal s)
		'Response.write JsCode(Str("alert('{1}');history.go(-1);",JsEncode(s)))
		Response.write JsCode(Str("alert('{1}');",JsEncode(s)))
	End Sub

	' 服务器端输出javascript弹出消息框并转到URL (来自Easp2.2)
	Public Sub AlertUrl(ByVal s, ByVal u)
		Response.write JsCode(Str("alert('{1}');location.href='{2}';",Array(JsEncode(s),u)))
	End Sub

	' 服务器端输出javascript确认消息框并根据选择转到URL (来自Easp2.2)
	Public Sub ConfirmUrl(ByVal s, ByVal t, ByVal f)
		Response.write JsCode(Str("location.href=confirm('{1}')?'{2}':'{3}';",Array(JsEncode(s),t,f)))
	End Sub
	
'================================================================================
'== Private
'================================================================================

'--------------------------------------------------------------------------------
'-- Function
'--------------------------------------------------------------------------------
	
	'-------------------------
	' 基础功能类函数
	'-------------------------

	' 取文件夹绝对路径 (参考Easp2.2 fso)
	' 调用: AbsPath_(路径[支持通配符*和?])
	' 返回: 物理路径
	' 说明: 地址第一个符号是/, 则位置从根目录开始(即:绝对地址), 不带则位置从当前文件所在文件夹开始(即:相对地址)
	Private Function AbsPath_(ByVal pPath)
		If Len(Trim(pPath))=0 Then AbsPath_ = "" : Exit Function
		If Mid(pPath,2,1)<>":" Then
			If IsWildcards_(pPath) Then
				pPath = Replace(pPath,"*","^$^")
				pPath = Replace(pPath,"?","^#^")
				pPath = Server.MapPath(pPath)
				pPath = Replace(pPath,"^$^","*")
				pPath = Replace(pPath,"^#^","?")
			Else
				pPath = Server.MapPath(pPath)
			End If
		End If
		If Right(pPath,1)="\" Then pPath = Left(pPath,Len(pPath)-1)
		AbsPath_ = pPath
	End Function

	' 内容编码转换
	Private Function CharsetTo_(ByVal pStr, ByVal pCharset) 
		dim loStream
		set loStream = Server.CreateObject("Adodb.Stream")
		With loStream
			.Type = 1 ' 1 二进制，2 文本
			.Mode = 3
			.Open
			.Write pStr	' Write 写二进制 , WriteText 写文本
			.Position = 0
			.Type = 2
			.Charset = pCharset
			CharsetTo_ = .ReadText	' 返回经转换后的文本内容
			.Close
		End With
		set loStream = nothing
	End Function

	' 正则替换内容
	' 调用: RegReplace_(内容, 规则, 替换成的内容, 是否多行替换)
	' 返回: 被替换后的内容
	Private Function RegReplace_(ByVal pStr, ByVal pRule, Byval pResult, ByVal pIsMult)
		Dim lvStr : lvStr = pStr
		If Has(pStr) Then
			If pIsMult=1 Then coRegExp.Multiline = True
			coRegExp.Pattern = pRule
			lvStr = coRegExp.Replace(lvStr, pResult)
			If pIsMult=1 Then coRegExp.Multiline = False
			coRegExp.Pattern = ""
		End If
		RegReplace_ = lvStr
	End Function

	' 格式化字符串（首下标为1）(来自Easp2.2)
	Private Function Str(ByVal s, ByVal v)
		Str = FormatString(s, v, 1)
	End Function

	' 格式化字符串（首下标为0）(来自Easp2.2)
	Private Function Format(ByVal s, ByVal v)
		Format = FormatString(s, v, 0)
	End Function

	' 格式化字串中的特殊字符(如jsencode编译后的代码),转为正常字符 (来自Easp2.2)
	Private Function FormatString(ByVal s, ByRef v, ByVal t)
		Dim i,n,k
		s = Replace(s,"\\",Chr(0))
		s = Replace(s,"\{",Chr(1))
		Select Case VarType(v)
			'数组
			Case 8192,8194,8204,8209
				For i = 0 To Ubound(v)
					s = FormatReplace(s,i+t,v(i))
				Next
			'对象
			Case 9
				Select Case TypeName(v)
					'记录集
					Case "Recordset"
						For i = 0 To v.Fields.Count - 1
							s = FormatReplace(s,i+t,v(i))
							s = FormatReplace(s,v.Fields.Item(i+t).Name,v(i))
						Next
					'字典
					Case "Dictionary"
						For Each k In v
							s = FormatReplace(s,k,v(k))
						Next
					'Easp List
					'Case "EasyAsp_List"
					'	For i = 0 To v.End
					'		s = FormatReplace(s,i+t,v(i))
					'		s = FormatReplace(s,v.IndexHash(i),v(i))
					'	Next
					'正则搜索子集合
					Case "ISubMatches", "SubMatches"
						For i = 0 To v.Count - 1
							s = FormatReplace(s,i+t,v(i))
						Next
				End Select
			'字符串
			Case 8
				Select Case TypeName(v)
					'正则搜索集合
					Case "IMatch2", "Match"
						s = FormatReplace(s,t,v.Value)
						For i = 0 To v.SubMatches.Count - 1
							s = FormatReplace(s,i+t+1,v.SubMatches(i))
						Next
					'字符串
					Case Else
						s = FormatReplace(s,t,v)
				End Select
			Case Else
				s = FormatReplace(s,t,v)
		End Select
		s = Replace(s,Chr(1),"{")
		FormatString = Replace(s,Chr(0),"\")
	End Function

	' 格式化Format内标签参数 (来自Easp2.2)
	Private Function FormatReplace(ByVal s, ByVal t, ByVal v)
		Dim tmp,rule,ru,kind,matches,match
		v = IfHas(v,"")
		rule = "\{" & t & "(:((N[,\(%]?(\d+)?)|(D[^\}]+)|(E[^\}]+)|U|L|\d+([^\}]+)?))\}"
		If Test(s,rule) Then
			Set matches = RegMatch(s,rule)
			For Each match In matches
				kind = RegReplace(match.Value, rule, "$2")
				ru = match.Value
				Select Case Left(kind,1)
					'截取字符串
					Case "1","2","3","4","5","6","7","8","9"
						s = Replace(s, ru, CutStr(v,regReplace(kind,"^(\d+)(.+)?$","$1:$2")))
					'数字
					Case "N"
						If isNumeric(v) Then
							Dim style,group,parens,percent,deci
							style = RegReplace(kind,"^N([,\(%])?(\d+)?$","$1")
							If style = "," Then group = -1
							If style = "(" Then parens = -1
							If style = "%" Then percent = -1
							deci = RegReplace(kind,"^N([,\(%])?(\d+)?$","$2")
							If IsNull(style) And IsNull(deci) Then
								s = Replace(s, ru, IIF(Instr(Cstr(v),".")>0 And v<1,"0" & v,v),1,-1,1)
							Else
								deci = IfHas(deci,-1)
								If percent Then
									s = Replace(s, ru, FormatNumber(v*100,deci,-1) & "%",1,-1,1)
								Else
									s = Replace(s, ru, FormatNumber(v,deci,-1,parens,group),1,-1,1)
								End If
							End If
						End If
					'日期
					Case "D"
						If isDate(v) Then
							s = Replace(s, ru, DateTime(v,Mid(kind,2)),1,-1,1)
						End If
					'转大写
					Case "U"
						s = Replace(s, ru, UCase(v),1,-1,1)
					'转小写
					Case "L"
						s = Replace(s, ru, LCase(v),1,-1,1)
					'表达式
					Case "E"
						tmp = Replace(Mid(kind,2),"%s", "v")
						tmp = Eval(tmp)
						s = Replace(s, ru, tmp,1,-1,1)
				End Select
			Next
		Else
			s = Replace(s,"{" & t & "}",v,1,-1,1)
		End If
		FormatReplace = s
	End Function

	' 取指定长度的随机字符串 (来自Easp2.2)
	Private Function RandStr_(ByVal length, ByVal allowStr)
		Dim i
		If IsNull(allowStr) Then allowStr = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
		For i = 1 To length
			Randomize(Timer) : RandStr_ = RandStr_ & Mid(allowStr, Int(Len(allowStr) * Rnd + 1), 1)
		Next
	End Function

	' 内部多参数处理 (来自Easp2.2)
	Private Function MParam(ByVal s)
		Dim arr(1),t : t = Instr(s,":")
		If t > 0 Then
			arr(0) = Left(s,t-1) : arr(1) = Mid(s,t+1)
		Else
			arr(0) = s : arr(1) = ""
		End If
		MParam = arr
	End Function

	'-------------------------
	' 验证类函数(Is打头)
	'-------------------------
	
	' 路径是否包含通配符 (参考Easp2.2 fso)
	' 调用: IsWildcards_(路径)
	' 返回: true 包含, false 不包含
	Private Function IsWildcards_(ByVal pPath)
		IsWildcards_ = False
		If Instr(pPath, "*")>0 Or Instr(pPath, "?")>0 Then IsWildcards_ = True
	End Function

	'-------------------------
	' 系统取值类函数(Get打头)
	'-------------------------

	' 取得特定字符左边或右边的字符串 (来自Easp2.2)
	Private Function GetStr_(ByVal s, ByVal m, ByVal t)
		Dim n : n = Instr(s, m)
		If n>0 Then
			If t = 0 Then
				GetStr_ = Left(s, n-1)
			ElseIf t=1 Then
				GetStr_ = Mid(s, n + Len(m))
			End If
		Else
			GetStr_ = s
		End If
	End Function

	' 取得指定目录下的所有文件名列表
	' 调用: GetFolderFiles_(目录路径[支持相对与绝对])
	' 返回: 文件列表用,号分隔, 如: ebody.fso.asp,ebody.db.asp,ebody.tpl.asp
	Private Function GetFolderFiles_(ByVal pFolderPath)
		Dim loFolder, loFile, loFiles, lvFileNames
		If pFolderPath = "" Then pFolderPath = "."
		If IsFolder(pFolderPath) Then
			pFolderPath = AbsPath_(pFolderPath)
			Set loFolder = coFSO.GetFolder(pFolderPath)
			Set loFiles = loFolder.Files
			' 循环读取文件, 组构文件名清单
			For Each loFile in loFiles
				If IsEmpty(lvFileNames) Then
					lvFileNames = loFile.name
				Else
					lvFileNames = lvFileNames & "," & loFile.name
				End If
			Next
			' 清空对像
			Set loFolder = Nothing
			Set loFiles = Nothing
			' 返回
			GetFolderFiles_ = lvFileNames
		End If
		' 捕捉错误
		ErrLog_("GetFolderFiles_")
	End Function	

	' 将文件名列表转换为类名列表
	' 调用: GetClassList_(文件列表)
	' 返回: 转换为类库变量的清单字串, 如: tpl,db,fso,act
	' 说明: 只取文件名以前后两个.号为分隔中间的字符作为类名, 如:ebody.tpl.asp, 则相应的类名为tpl
	Private Function GetClassList_(ByVal pFileList)
		Dim lvFileNames, lvFileName, lvClassList, lvTempStr
		lvFileNames = Split(pFileList, ",")
		For Each lvFileName In lvFileNames
			' 取文件名两个.的中间一段做为类名	
			If Left(lvFileName, Len(cvGlobalClassName)+1)=cvGlobalClassName & "." Then
				lvTempStr = Left(lvFileName, InstrRev(lvFileName, ".")-1)
				If Len(lvClassList) = 0 Then
					lvClassList = Mid(lvTempStr, InStr(lvTempStr, ".")+1)
				Else
					lvClassList = Mid(lvTempStr, InStr(lvTempStr, ".")+1) & "," & lvClassList
				End If
			End If
		Next
		' 返回
		GetClassList_ = lvClassList
	End Function

	' 读取指定文本文件内容
	' 调用: GetFile_(文件路径[支持相对与绝对])
	' 返回: 文件字符串内容
	' 样例: GetFile_("abc/Readme.txt")
	' 说明: 读取当前文件内容, 通过数据流控件(ADODB.Stream)读取或FSO读取, 只能读取文本文件
	Private Function GetFile_(ByVal pFilePath)
		Dim lvFileData
		' 验证文件是否存在
		If Not IsFile(pFilePath) Then Exit Function
		' 1. 先以流方式读取
		lvFileData = GetFileByStream_(pFilePath, 2)
		' 判断是否成功
		If Err.Number <> 0 Then
			Err.Clear()
			' 2. 不成功,再以Fso方式读取
			lvFileData = GetFileByFSO_(pFilePath)
		End If
		' 处理文件BOM
		If LCase(cvCharset) = "utf-8" Then
			Select Case LCase(cvBom)
			Case "keep"
				' 什么都不做
			Case "remove"
				If Test(lvContent, "^\uFEFF") Then
					lvFileData = RegReplace(lvFileData, "^\uFEFF", "")
				End If
			Case "add"
				If Not Test(lvContent, "^\uFEFF") Then
					lvFileData = Chrw(&hFEFF) & lvFileData
				End If
			End Select
		End If
		' 返回
		GetFile_ = lvFileData
	End Function

	' 无限级读取文件内容, 同时也读取所有包含文件的内容
	' 调用: GetFileAll_(文件路径[支持相对与绝对])
	' 返回: 所有Include中的文件的内容
	' 说明: 此方法支持带有<inclued>标签的文件内容读取
	Private Function GetFileAll_(ByVal pFilePath)
		Dim lvContent, lvRule, lvIncFile, lvIncStr
		Dim loMatch, loMatches
		' 读取当前文件的内容
		pFilePath = AbsPath_(pFilePath)
		lvContent = GetFile_(pFilePath)
		If IsNull(lvContent) Then Exit Function
		' 替换为特定标签
		lvContent = RegReplace(lvContent, "<% *?@.*?%"&">", "")
		lvContent = RegReplace(lvContent, "(<%[^>]+?)(option +?explicit)([^>]*?%"&">)", "$1'$2$3")
		' 定义搜寻包含文件的规则
		lvRule = "<!-- *?#include +?(file|virtual) *?= *?""??([^"":?*\f\n\r\t\v]+?)""?? *?-->"		
		' 验证内容是否有效, 并进行读取
		If RegTest(lvContent, lvRule) Then
			Set loMatches = RegMatch(lvContent, lvRule)
			For Each loMatch In loMatches
				If LCase(loMatch.SubMatches(0)) = "virtual" Then
					lvIncFile = loMatch.SubMatches(1)
				Else
					lvIncFile = Mid(pFilePath, 1, InstrRev(pFilePath,IIF(Instr(pFilePath, ":")>0, "\", "/"))) & loMatch.SubMatches(1)
				End If
				lvIncStr = GetFileAll_(lvIncFile)
				lvContent = Replace(lvContent, loMatch, lvIncStr)
			Next
			Set loMatches = Nothing
		End If
		GetFileAll_ = lvContent
	End Function

	' 读取文件内容方法1 (使用数据流控件ADODB.Stream读取)
	' 调用: GetFileByStream_(文件路径[支持相对与绝对], 读取类型[1 二进制模式/2 文本模式])
	' 返回: 依据读取类型返回, pReadType=1 返回二进制内容, pReadType=2 返回文本内容
	' 说明: 只读取当前文件内容, 不支持多级include读取
	Private Function GetFileByStream_(ByVal pFilePath, ByVal pReadType)
		Dim lvFileData, loStream
		pFilePath = AbsPath_(pFilePath)
		Set loStream = Server.CreateObject("ADODB.Stream")
		' 开始读取
		With loStream
			.Type = pReadType	' 读取模式, 1 二进制 / 2 文本
			.Mode = 3			' 可读可写	
			.Open
			.LoadFromFile pFilePath
			.Charset = cvCharset
			.Position = 2		' 或为2
			If pReadType = 1 Then
				lvFileData = .Read		' 以二进制模式读取(返回二进制内容)
			Else
				lvFileData = .ReadText	' 以文本模式读取(返回文本内容)
			End If
			.Close
		End With
		Set loStream = Nothing
		' 返回
		GetFileByStream_ = lvFileData		
	End Function

	' 读取文件内容方法2 (使用文件系统控件Fso读取)
	' 调用: GetFileByFSO_(文件路径[支持相对与绝对])
	' 返回: 文件字符串内容
	' 说明: 只能读取文本文件(txt), 不支持大字符集文件
	Private Function GetFileByFSO_(ByVal pFilePath)
		Dim loFile
		pFilePath = AbsPath_(pFilePath)
		' 打开并读取文件内容, 直到文件结尾, 则返回
		Set loFile = coFSO.OpenTextFile(pFilePath, 1, False, -1)
		If loFile.AtEndOfStream = false Then GetFileByFSO_ = loFile.ReadAll
		' 关闭文件
		loFile.Close
		Set loFile=Nothing
	End Function

'--------------------------------------------------------------------------------
'-- Sub
'--------------------------------------------------------------------------------

	'-------------------------
	' 基础功能类过程
	'-------------------------

	' 释放指定清单对像的内存(关闭清单中指定的Ebody类实例对象)
	' 调用: CloseObjFromList_(ByVal pClassList)
	Private Sub CloseObjFromList_(ByVal pClassList)
		Dim lvClassList, i
		lvClassList = Split(pClassList, ",")
		For i = Ubound(lvClassList) To 0 Step -1
			' 验证对像是否存在
			If Eval("IsObject(" & lvClassList(i) & ")") Then
				' 动态关闭对像
				Execute "Set " & lvClassList(i) & " = Nothing"
			End If
		Next
	End Sub

	' 记录错误
	' 说明: 记录系统出现的最后一次错误信息
	Private Sub ErrLog_(pMsg)
			cvErrCode = Err.Number
			cvErrMsg = pMsg
	End Sub

	' 注册所有扩展子类,并实例化扩展基类(ext)
	' 说明: 实例化此基类后,后继就可使用ext.的方法直接引用扩展子类的方法
	'		其中将提取扩展类目录下的所有类文件,进行注册
	Private Sub RegisterExtClass_()
		' 得到扩展类库目录文件名清单
		Dim lvExtFileList
		lvExtFileList = GetFolderFiles_(cvExtClassPath)
		' 将文件清单转换为类库变量名清单
		Dim lvExtClassList
		lvExtClassList = GetClassList_(lvExtFileList)
		' 动态组构类库变量清单
		Dim lvVarList
		lvVarList = "Public " & lvExtClassList
		' 动态组构类库中的close方法
		Dim lvSubCode
		lvSubCode = ": Public Sub Close(pObj) Execute(""Set "" & pObj & ""=Nothing"") End Sub :"
		' 动态创建扩展类实例
		If Not IsEmpty(lvExtClassList) Then
			' 注册: 组构Ext扩展类内容,加入扩展类对应的变量,将扩展类代码载入内存(动态创建扩展类中的变量)
			ExecuteGlobal("Class ebody_ext " & lvVarList & lvSubCode & " End Class")
			' 实例化扩展类(将实例化的扩展类动态映射到EXT变量中)
			Execute("Set ext = new ebody_ext")
		End If
	End Sub

	'-------------------------
	' 设置对像类过程(Set打头)
	'-------------------------

End Class
%>