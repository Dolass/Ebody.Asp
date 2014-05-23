<%
'################################################################################
'## ebody.http.asp
'## -----------------------------------------------------------------------------
'## 功能:	HTTP网页读写类
'## 版本:	1.1 Build 130324
'## 作者:	Tony
'## 日期:	2013/03/24
'## 说明:	Ebody基类
'################################################################################

Class ebody_http
	
'================================================================================
'== Variable
'================================================================================


'--------------------------------------------------------------------------------
'-- Public
'--------------------------------------------------------------------------------

	Public CharSet	' 字符编码
	Public Async	' 异步模式开关	
	Public User		' 访问目标URL用户名
	Public Password	' 访问目标URL密码
	Public Url		' 目标网络地址
	Public Method	' 发送方式 GET/POST
	Public Html		' 接收返回HTML内容
	Public Headers	' 响应文件头内容
	Public Body		' 响应文件正文内容
	Public Text		' 响应文件文本内容
	Public SaveRandom	' 保存文件是否随机命名
	Public ResolveTimeout	' 服务器解析超时（毫秒）
	Public ConnectTimeout	' 服务器连接超时（毫秒）
	Public SendTimeout		' 发送数据超时（毫秒）
	Public ReceiveTimeout	' 接受数据超时（毫秒）

'--------------------------------------------------------------------------------
'-- Private
'--------------------------------------------------------------------------------

	Private cvData	' 客户化数据
	Private cvUrl	' 请求地址
	Private coRh	' 请求头信息存储字典
	Private coList	' 组存储对像
	Private cvHtml	' Html临时存储


'================================================================================
'== Event
'================================================================================

	Private Sub Class_Initialize()
		'On Error Resume Next		
		CharSet = ""	' 编码默认为空，将自动获取编码		
		Async = False	' 异步模式关闭
		User = ""
		Password = ""		
		Html = ""
		Headers = ""
		Body = Empty
		Text = Empty
		SaveRandom = True
		ResolveTimeout = 20000
		ConnectTimeout = 20000
		SendTimeout = 300000
		ReceiveTimeout = 60000

		cvData = ""
		cvUrl = ""

		Set coRh = Server.CreateObject("Scripting.Dictionary")	' 请求头信息存储字典
		coList = Array()
	End Sub

	Private Sub Class_Terminate()
		'On Error Resume Next
		Set coRh = Nothing
	End Sub

'================================================================================
'== Public
'================================================================================


'--------------------------------------------------------------------------------
'-- Let Property
'--------------------------------------------------------------------------------

	' 设置要提交的数据
	Public Property Let Data(ByVal p)
		cvData = p
	End Property

	' 设置单项请求头信息
	Public Property Let RequestHeader(ByVal n, ByVal v)
		n = Replace(n,"-","_")
		coRh(n) = v
	End Property


'--------------------------------------------------------------------------------
'-- Get Property
'--------------------------------------------------------------------------------

	' 获取单项请求头信息
	Public Property Get RequestHeader(ByVal n)
		If Ebody.Has(n) Then
			RequestHeader = coRh(n)
		Else
			i = 0
			For Each lvRh In coRh
				coList(i) = lvRh & ":" & coRh(lvRh)
			Next
			' 组构数组值为字串
			RequestHeader = Join(coList,vbCrLf)
		End If
	End Property


'--------------------------------------------------------------------------------
'-- Function
'--------------------------------------------------------------------------------

	'-------------------------
	' 基础功能类函数
	'-------------------------

	' 建新Ebody远程文件操作类实例
	Public Function [New]()
		Set [New] = New ebody_http
	End Function

	' Get模式取远程页
	Public Function [Get](ByVal pUrl)
		[Get] = GetData(pUrl, "GET", Async, cvData, User, Password)
	End Function

	' Post模式取远程页
	Public Function Post(ByVal pUrl)
		Post = GetData(pUrl, "POST", Async, cvData, User, Password)
	End Function

	' 属性配置模式下打开连接远程
	Public Function [Open]
		[Open] = GetData(Url, Method, Async, cvData, User, Password)
	End Function

	' 按标签查找字符串
	' 在返回的响应内容中按条件查找指定字符串，找到后并返回匹配内容
	' pTagStart - 要截取的部分的开头
	' pTagEnd   - 要截取的部分的结尾
	' pTagSelf  - 结果是否包括tagStart和tagEnd
	'           (0或空:不包括,1:包括,2:只包括tagStart,3:只包括tagEnd)
	' cvHtml - 返回的响应内容，用Get或Post取得目标地址所返回的内容
	Public Function SubStr(ByVal pTagStart, ByVal pTagEnd, ByVal pTagSelf)	
		SubStr = SubStr_(cvHtml,pTagStart,pTagEnd,pTagSelf)
	End Function


	' 在返回的响应内容中按条件查找指定字符串，找到后并返回匹配内容
	Private Function SubStr_(ByVal pHtml, ByVal pTagStart, ByVal pTagEnd, ByVal pTagSelf)
		Dim lvPosA, lvPosB, lvFirst, lvBetween
		lvPosA = instr(1,pHtml,pTagStart,1)
		If lvPosA=0 Then SubStr_ = "源代码中不包括此开始标签" : Exit Function
		lvPosB = instr(lvPosA+Len(pTagStart),pHtml,pTagEnd,1) 
		If lvPosB=0 Then SubStr_ = "源代码中不包括此结束标签" : Exit Function
		Select Case pTagSelf
			Case 1
				lvFirst = lvPosA
				lvBetween = lvPosB+len(pTagEnd)-lvFirst
			Case 2
				lvFirst = lvPosA
				lvBetween = lvPosB-lvFirst
			Case 3
				lvFirst = lvPosA+len(pTagStart)
				lvBetween = lvPosB+len(pTagEnd)-lvFirst
			Case Else
				lvFirst = lvPosA+len(pTagStart)
				lvBetween = lvPosB-lvFirst
		End Select
		SubStr_ = Mid(pHtml,lvFirst,lvBetween)
	End Function

	' 按正则查找符合的第一个字符串
	' cvHtml - 返回的响应内容，用Get或Post取得目标地址所返回的内容
	Public Function Find(ByVal pRule)
		Find = Find_(cvHtml, pRule)
	End Function

	' 按正则查找符合的第一个字符串
	Private Function Find_(ByVal pHtml, ByVal pRule)
		If Ebody.Test(pHtml,pRule) Then Find_ = Ebody.RegReplace(pHtml,"([\s\S]*)("&pRule&")([\s\S]*)","$2")
	End Function


	' 按正则查找符合的第一个字符串，可按正则编组选择其中的一部分
	Public Function [Select](ByVal pRule, ByVal pPart)
		[Select] = Select_(cvHtml, pRule, pPart)
	End Function

	' 按正则查找符合的第一个字符串，可按正则编组选择其中的一部分
	Private Function Select_(ByVal pHtml, ByVal pRule, ByVal pPart)
		If Ebody.Test(pHtml,pRule) Then
			'$0匹配字符串本身
			pPart = Replace(pPart,"$0",Find_(pHtml,pRule))
			'按正则编组分别替换
			Select_ = Ebody.RegReplace(pHtml,"(?:[\s\S]*)(?:"&pRule&")(?:[\s\S]*)",pPart)
		End If
	End Function

	' 按正则查找符合的字符串组，返回数组
	Public Function Search(ByVal pRule)
		Search = Search_(cvHtml, pRule)
	End Function

	' 按正则查找符合的字符串组，返回数组
	Private Function Search_(ByVal pHtml, ByVal pRule)
		Dim matches,match,arr(),i : i = 0
		Set matches = Ebody.RegMatch(pHtml,pRule)
		ReDim arr(matches.Count-1)
		For Each match In matches
			arr(i) = match.Value
			i = i + 1
		Next
		Set matches = Nothing
		Search_ = arr
	End Function

	' 保存远程图片到本地
	Public Function SaveImgTo(ByVal p)
		SaveImgTo = SaveImgTo_(cvHtml,p)
	End Function

	' 保存远程图片到本地
	Private Function SaveImgTo_(ByVal s, ByVal p)
		Dim a,b, i, img, ht, tmp, src
		' 取得图片地址
		a = GetImg(s)
		b = GetImgTag(s)
		If Ebody.Has(a) Then
			Ebody.Use "Fso"
			For i = 0 To Ubound(a)
				' 生成随机字符
				If SaveRandom Then
					img = Ebody.DateTime(Now,"ymmddhhiiss"&Ebody.RandStr("5:0123456789")) & Mid(a(i),InstrRev(a(i),"."))
				Else
					img = Mid(a(i),InstrRev(a(i),"/")+1)
				End If

				' 创建新的http实例
				Set ht = Ebody.Http.New
				
				' 取得页面图片数据
				ht.Get a(i)
				
				' 生成本地图片文件
				tmp = Ebody.Fso.SaveAs(p & img, ht.Body)

				Set ht = Nothing
				If tmp Then
					'Response.write "b(i)=> " & Ebody.HtmlEncode(b(i))
					src = Ebody.RegReplace(b(i),"(<img\s[^>]*src\s*=\s*([""|']?))("&a(i)&")(\2[^>]*>)","$1"&p&img&"$4")
					'Response.write "src=> " & Ebody.HtmlEncode(src)
					s = Replace(s,b(i),src)
				End If
			Next
		End If
		SaveImgTo_ = s
	End Function


	' 获取文本中的图片地址存为一个数组
	Private Function GetImg(ByVal s)
		GetImg = GetImg_(s,0)
	End Function

	' 获取文本中的图像标签存为一个数组
	Private Function GetImgTag(ByVal s)
		GetImgTag = GetImg_(s,1)
	End Function

	' 取得图片地址
	Private Function GetImg_(ByVal s, ByVal t)
		Dim a(), img, Matches, m, i : i = 0
		ReDim a(-1)
		img = "<img([^>]+?)(/?)>"
		If Ebody.Has(s) And Ebody.RegTest(s, img) Then
			'取消所有的换行和缩进
			s = Replace(s, vbCrLf, " ")
			s = Replace(s, vbTab, " ")
			'正则匹配所有的img标签
			Set Matches = Ebody.RegMatch(s, "(<img\s[^>]*src\s*=\s*([""|']?))([^""'>]+)(\2[^>]*>)")
			'取出每个img的src存入数组
			For Each m In Matches
				ReDim Preserve a(i)
				a(i) = Ebody.IIF(t=0,m.SubMatches(2),m.value)
				i = i + 1
			Next
		End If
		GetImg_ = a
	End Function

	'-------------------------
	' 验证类函数(Is打头)
	'-------------------------

	'-------------------------
	' 系统取值类函数(Get打头)
	'-------------------------

	' 获取远程页完整参数模式
	Public Function GetData(ByVal pUrl, ByVal pMethod, ByVal pAsync, ByVal pData, ByVal pUser, ByVal pPassword)
		Dim loHTTP
		'建立XMLHttp对象
		If Ebody.isInstall("MSXML2.serverXMLHTTP") Then
			Set loHTTP = Server.CreateObject("MSXML2.serverXMLHTTP")
		ElseIf Ebody.isInstall("MSXML2.XMLHTTP") Then
			Set loHTTP = Server.CreateObject("MSXML2.XMLHTTP")
		ElseIf Ebody.isInstall("Microsoft.XMLHTTP") Then
			Set loHTTP = Server.CreateObject("Microsoft.XMLHTTP")
		Else
			'Ebody.Error.Raise 47
			Exit Function
		End If

		' 设置超时时间
		loHTTP.SetTimeOuts ResolveTimeout, ConnectTimeout, SendTimeout, ReceiveTimeout

		' 抓取地址
		If Ebody.IsNull(pUrl) Then Exit Function

		' 通过URL临时指定编码
		If Ebody.Test(pUrl,"^[\w\d-]+>https?://") Then
			CharSet = Ebody.CLeft(pUrl, ">")
			pUrl = Ebody.CRight(pUrl, ">")
		End If
		cvUrl = pUrl

		' 方法：POST或GET
		pMethod = Ebody.IIF(Ebody.Has(pMethod), UCase(pMethod), "GET")

		' 异步
		If Ebody.IsNull(pAsync) Then pAsync = False

		' 构造Get传数据的URL
		If pMethod = "GET" And Ebody.Has(pData) Then pUrl = pUrl & Ebody.IIF(Instr(pUrl,"?")>0, "&", "?") & Serialize_(pData)
		
		' 打开远程页
		If Ebody.Has(pUser) Then
			' 如果有用户名和密码
			loHTTP.open pMethod, pUrl, pAsync, pUser, pPassword
		Else
			' 匿名
			loHTTP.open pMethod, pUrl, pAsync
		End If

		If pMethod = "POST" Then
			If Not coRh.Exists("Content_Type") Then
				coRh("Content_Type") = "application/x-www-form-urlencoded"
			End If
			
			' 设置RequestHeader
			SetHeaderTo loHTTP

			' 有发送的数据
			loHTTP.send Serialize_(pData)
		Else
			' 设置RequestHeader
			SetHeaderTo loHTTP

			' 无数据发送
			loHTTP.send
		End If

		' 检测返回数据
		If loHTTP.readyState <> 4 Then
			' 无响应退出
			GetData = "error:server is down"
			Set loHTTP = Nothing
			'Ebody.Error.Raise 46
			Exit Function
		ElseIf loHTTP.Status = 200 Then
			Headers = loHTTP.getAllResponseHeaders()
			Body = loHTTP.responseBody
			Text = loHTTP.responseText
			If Ebody.IsNull(CharSet) Then
				' 从Header中提取编码信息
				If Ebody.Test(Headers,"charset=([\w-]+)") Then
					' 一般文档
					CharSet = Ebody.RegReplace(Headers,"([\s\S]+)charset=([\w-]+)([\s\S]+)","$2")
				
				ElseIf Ebody.Test(Headers,"Content-Type: ?text/xml") Then
					' 如果是Xml文档，从文档中提取编码信息
					CharSet = Ebody.RegReplace(Text,"^<\?xml\s+[^>]+encoding\s*=\s*""([^""]+)""[^>]*\?>([\s\S]+)","$1")
				
				ElseIf Ebody.Test(Text,"<meta\s+http-equiv\s*=\s*[""']?content-type[""']?\s+content\s*=\s*[""']?[^>]+charset\s*=\s*([\w-]+)[^>]*>") Then
					' 从文件源码中提取编码
					CharSet = Ebody.RegReplace(Text,"([\s\S]+)<meta\s+http-equiv\s*=\s*[""']?content-type[""']?\s+content\s*=\s*[""']?[^>]+charset\s*=\s*([\w-]+)[^>]*>([\s\S]+)","$2")
				End If

				' 如果无法获取远程页的编码则继承Ebody的编码设置
				If Ebody.IsNull(CharSet) Then CharSet = Ebody.CharSet
			End If
			GetData = Ebody.CharsetTo(Body, CharSet)
		Else
			GetData = "error:" & loHTTP.Status & " " & loHTTP.StatusText
		End If
		Set loHTTP = Nothing
		cvHtml = GetData
		Html = cvHtml
	End Function

'--------------------------------------------------------------------------------
'-- Sub
'--------------------------------------------------------------------------------
	
	'-------------------------
	' 基础功能类过程
	'-------------------------
	
	' 启用Ajax代理
	Public Sub AjaxAgent()
		Ebody.NoCache()
		Dim u, qs, qskey, qf, qfkey, m
		' 取得目标地址
		u = Ebody.Get("ebodyurl")
		If Ebody.IsNull(u) Then Response.write "error:Invalid URL"
		If Instr(u,"?")>0 Then
			qs = "&" & Ebody.CRight(u,"?")
			u = Ebody.CLeft(u,"?")
		End If

		' 传url参数
		If Request.QueryString()<>"" Then
			For Each qskey In Request.QueryString
				If qskey<>"ebodyurl" Then qs = qs & "&" & qskey & "=" & Request.QueryString(qskey)
			Next
		End If
		u = u & Ebody.IfThen(Ebody.Has(qs),"?" & Mid(qs,2))

		' 如果是Post则同时传Form数据
		m = Request.ServerVariables("REQUEST_METHOD")
		If m = "POST" Then
			If Request.Form()<>"" Then
				For Each qfkey In Request.Form
					qf = qf & "&" & qfkey & "=" & Request.Form(qfkey)
				Next
				Data = Mid(qf,2)
			End If
			Response.write Post(u)
		Else
			Response.write [Get](u)
		End If
	End Sub

	'-------------------------
	' 设置对像类过程(Set/Remove打头)
	'-------------------------

	' 设置请求头信息
	Public Sub SetHeader(ByVal p)
		Dim i,n,v
		If isArray(p) Then
			For i = 0 To Ubound(p)
				n = Replace(Ebody.Cleft(p(i),":"),"-","_")
				v = Ebody.CRight(p(i),":")
				coRh(n) = v
			Next
		Else
			n = Replace(Ebody.Cleft(p,":"),"-","_")
			v = Ebody.CRight(p,":")
			coRh(n) = v
		End If
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

	' url参数化序列化
	Private Function Serialize_(ByVal p)
		Dim tmp, i, n, v : tmp = ""
		If Ebody.IsNull(p) Then Exit Function
		If isArray(p) Then
			For i = 0 To Ubound(p)
				n = Ebody.CLeft(p(i),":")
				v = Ebody.CRight(p(i),":")
				tmp = tmp & "&" & n & "=" & Server.URLEncode(v)
			Next
			If Len(tmp)>1 Then tmp = Mid(tmp,2)
			Serialize_ = tmp
		Else
			Serialize_ = p
		End If
	End Function

	'-------------------------
	' 验证类函数(Is打头)
	'-------------------------

	'-------------------------
	' 系统取值类函数(Get打头)
	'-------------------------

'--------------------------------------------------------------------------------
'-- Sub
'--------------------------------------------------------------------------------

	'-------------------------
	' 基础功能类过程
	'-------------------------


	'-------------------------
	' 设置对像类过程(Set打头)
	'-------------------------
	
	' 传入RequestHeader
	Private Sub SetHeaderTo(ByRef pObj)
		Dim maps, key
'		Set maps = o_rh.Maps
		For Each key In coRh
			If Not isNumeric(key) Then
				pObj.setRequestHeader Replace(key,"_","-"), coRh(key)
			End If
		Next
		Set maps = Nothing
	End Sub

End Class
%>