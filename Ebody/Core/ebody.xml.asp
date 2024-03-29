<%
'################################################################################
'## ebody.xml.asp
'## -----------------------------------------------------------------------------
'## 功能:	xml文本控制类
'## 版本:	1.1 Build 130324
'## 作者:	Tony
'## 日期:	2013/03/24
'## 说明:	Ebody基类
'################################################################################

Class ebody_xml
	
'================================================================================
'== Variable
'================================================================================


'--------------------------------------------------------------------------------
'-- Public
'--------------------------------------------------------------------------------

	Public Dom		' DOMDocument对像
	Public Doc		' DOMDocument对像
	Public IsOpen

'--------------------------------------------------------------------------------
'-- Private
'--------------------------------------------------------------------------------

	Private s_filePath
	Private s_xsltPath

'================================================================================
'== Event
'================================================================================

	Private Sub Class_Initialize()
		'On Error Resume Next
		' 初始化
		IsOpen = False
		s_filePath = ""
		s_xsltPath = ""
	End Sub

	Private Sub Class_Terminate()
		'On Error Resume Next
		'释放Document
		If IsObject(Doc) Then Set Doc = Nothing
		If IsObject(Dom) Then Set Dom = Nothing
	End Sub

'================================================================================
'== Public
'================================================================================


'--------------------------------------------------------------------------------
'-- Let Property
'--------------------------------------------------------------------------------

	' 设置样式文件
	Public Property Let XSLT(ByVal x)
		Dim pi
		Set pi = Dom.CreateProcessingInstruction("xml-stylesheet", "type=""text/xsl"" href=""" & x & """")
		If Dom.ChildNodes(1).BaseName<>"xml-stylesheet" Then
			If Dom.FirstChild.BaseName<>"xml" Then
				Dom.InsertBefore pi, Dom.FirstChild
			Else
				Dom.InsertBefore pi, Dom.ChildNodes(1)
			End If
		Else
			Dom.ReplaceChild pi, Dom.ChildNodes(1)
		End If
		s_xsltPath = x
	End Property
	

'--------------------------------------------------------------------------------
'-- Get Property
'--------------------------------------------------------------------------------

	' 取得样式文件
	Public Property Get XSLT
		XSLT = s_xsltPath
	End Property

'--------------------------------------------------------------------------------
'-- Function
'--------------------------------------------------------------------------------

	'-------------------------
	' 基础功能类函数
	'-------------------------

	' 打开一个已经存在的XML文件,将内容载入到内存,返回打开状态
	Public Function Open(byVal f)
		Open = False
		If Ebody.IsNull(f) Then Exit Function
		Set Dom = NewDom()
		'转换为绝对路径
		f = Ebody.MapPath(f)
		'读取文件
		Dom.load f
		'存路径（用于保存）
		s_filePath = f
		If Not IsErr Then
			'设置根元素
			Set Doc = NewNode(Dom.documentElement)
			Open = True
			IsOpen = True
		Else
			Set Dom = Nothing
		End If
	End Function
	
	' 建立新的Ebody Node对象
	Public Function NewNode(ByVal o)
		Set NewNode = New Ebody_Xml_Node
		NewNode.Dom = o
	End Function
		
	' 建立新的根对象
	Public Function Root
		Set Root = NewNode(Dom)
	End Function
	
	' 建立新的Ebody Xml对象
	Public Function [New]()
		Set [New] = New Ebody_Xml
	End Function
		
	' 根据TagName取对象
	Public Default Function Find(ByVal t)
		Dim o,s
		'如果是Html代码片断
		If Ebody.Test(t,"^<[\s\S]+>$") Then
'			Dim n,a,v,r
'			r = "^<([^\s>]+)\s([^>]+)>([\s\S]+)</\1>$"
'			n = Ebody.RegReplace(t,r,"$1")
'			a = Ebody.RegReplace(t,r,"$2")
'			v = Ebody.RegReplace(t,r,"$3")
		Else
			If Ebody.Test(t, "[, >\[@:]") Then
				'按简单表达式取元素
				Set o = Dom.selectNodes(Ebody_Xml_TransToXpath(t))
			Else
				'从标签取元素
				Set o = Dom.GetElementsByTagName(t)
			End If
		End If
		
		If o.Length = 0 Then
			' 如果没有
			Exit Function
			'Ebody.Error.Msg = "("&t&")"
			'Ebody.Error.Raise 98		
		ElseIf o.Length = 1 Then
			' 如果只有一个元素
			Set Find = NewNode(o(0))		
		Else
			'如果是元素集合
			Set Find = NewNode(o)
		End If
	End Function

'	Function RenderToNode(ByVal s, ByRef node)
'		Dim n,a,v,r
'		r = "^<([^\s>]+)\s([^>]+)>([\s\S]+)</\1>$"
'		n = Ebody.RegReplace(t,r,"$1")
'		a = Ebody.RegReplace(t,r,"$2")
'		v = Ebody.RegReplace(t,r,"$3")
'		Ebody.WNH n
'		Ebody.WNH a
'		Ebody.WNH v
'	End Function

	' XPath取对象集合
	Public Function [Select](ByVal p)
		Set [Select] = NewNode(Dom.selectNodes(p))
	End Function

	' XPath取单个对象
	Public Function Sel(ByVal p)
		Set Sel = NewNode(Dom.selectSingleNode(p))
	End Function

	' 新建一个节点
	Public Function Create(ByVal n, ByVal v)
		Dim o,p,cd

		'类型可在名称中用空格隔开，例："mytag cdata", " comment"
		If Instr(n," ")>0 Then
			cd = LCase(Ebody.CRight(n," "))
			n = Ebody.CLeft(n," ")
		End If
		'创建注释节点
		If cd="comment" Then
			Set o = Dom.CreateComment(v)
		Else
			'创建节点
			Set o = Dom.CreateElement(n)
			If cd = "cdata" Then
				'创建CDATASection节点
				Set p = Dom.CreateCDATASection(v)
			Else
				'创建文本节点
				Set p = Dom.CreateTextNode(v)
			End If
			'追加到节点
			o.AppendChild(p)
		End If
		'返回新建的Node对象
		Set Create = NewNode(o)
		Set o = Nothing
		Set p = Nothing
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

	' 关闭文件
	Public Sub Close()
		Set Doc = Nothing
		Set Dom = Nothing
		s_filePath = ""
		IsOpen = False
	End Sub

	' 保存文件
	Public Sub [Save]()
		If IsOpen Then
			Dom.Save(s_filePath)
		Else
			Exit Sub
			'Ebody.Error.Msg = "（文档未处于打开状态）"
			'Ebody.Error.Raise 99
		End If
	End Sub

	' 另存为(Load进来的只能用此方法保存)
	' 可以在保存的文件名后加如 >gbk 来指定保存的编码，如果不指定则：有编码申明的不改变，无编码申明的采用Ebody.Charset值作编码
	Public Sub SaveAs(ByVal p)
		Dim ch,cha,pi
		If Instr(p,">")>0 Then
			ch = Ebody.CRight(p,">")
			p = Ebody.CLeft(p,">")
		End If
		cha = Ebody.IfHas(ch,Ebody.CharSet)
		p = Ebody.MapPath(p)

		' 如果没有文档类型申明就加上
		Set pi = Dom.CreateProcessingInstruction("xml", "version=""1.0"" encoding=""" & cha & """")
		If Dom.FirstChild.BaseName<>"xml" Then
			Dom.InsertBefore pi, Dom.FirstChild
		Else
			If Ebody.Has(ch) Then Dom.ReplaceChild pi, Dom.FirstChild
		End If
		Dom.Save(p)
		Set pi = Nothing
	End Sub

	' 用XSLT将XML转换为XHTML文档
	Public Sub SaveAsXHTML(ByVal p, ByVal xsl)
		Dim x,f : Set x = [New]
		If Ebody.Test(xsl,"^([\w\d-]+>)?https?://") Then
			x.Load xsl
		Else
			x.Open xsl
		End If
		f = Dom.TransformNode(x.Dom)
		Ebody.Use "Fso"
		Ebody.Fso.CreateTextFile p, f
		Set x = Nothing
	End Sub


	' 从文本或者URL载入XML结构数据
	Public Sub Load(ByVal s)
		If Ebody.IsNull(s) Then Exit Sub
		Dim str
		' 如果是外部网址则用Http取回,如要指定编码可加在http前，例：gbk>http://....
		If Ebody.Test(s,"^([\w\d-]+>)?https?://") Then
			Ebody.Use "Http"
			Dim h : Set h = Ebody.Http.New
			str = h.Get(s)
			Set h = Nothing
		Else
			str = s
		End If
		Set Dom = NewDom()
		'从文本加载
		Dom.loadXML(str)
		'设置根元素
		If Not IsErr Then
			Set Doc = NewNode(Dom.documentElement)
		Else
			Set Dom = Nothing
		End If
	End Sub

	'-------------------------
	' 设置对像类过程(Set/Remove打头)
	'-------------------------

'================================================================================
'== Private
'================================================================================


'--------------------------------------------------------------------------------
'-- Function
'--------------------------------------------------------------------------------

	'-------------------------
	' 基础功能类函数
	'-------------------------

	' 创建新的Xml对象
	Private Function NewDom()
		Dim o
		If Ebody.IsInstall("MSXML2.DOMDocument") Then
		' msxml ver 3
			Set o = Server.CreateObject("MSXML2.DOMDocument")
		ElseIf Ebody.IsInstall("Microsoft.XMLDOM") Then
		' msxml ver 2
			Set o = Server.CreateObject("Microsoft.XMLDOM")
		End If
		' 保留空格
		o.PreserveWhiteSpace = True
		' 异步
		o.Async = False
		' 使用Xpath表达式
		o.SetProperty "SelectionLanguage", "XPath"
		Set NewDom = o
		Set o = Nothing
	End Function

	'-------------------------
	' 验证类函数(Is打头)
	'-------------------------

	' 检查并打印错误信息
	Private Function IsErr()
		Dim s
		IsErr = False
	  If Dom.ParseError.Errorcode<>0 Then
			With Dom.ParseError
				s = s & "	<ul class=""dev"">" & vbCrLf
				s = s & "		<li class=""info"">以下信息针对开发者：</li>" & vbCrLf
				s = s & "		<li>错误代码：0x" & Hex(.Errorcode) & "</li>" & vbCrLf
				If Ebody.Has(.Reason) Then s = s & "		<li>错误原因：" & .Reason & "</li>" & vbCrLf
				If Ebody.Has(.Url) Then s = s & "		<li>错误来源：" & .Url & "</li>" & vbCrLf
				If Ebody.Has(.Line) And .Line<>0 Then s = s & "		<li>错误行号：" & .Line & "</li>" & vbCrLf
				If Ebody.Has(.Filepos) And .Filepos<>0 Then s = s & "		<li>错误位置：" & .Filepos & "</li>" & vbCrLf
				If Ebody.Has(.SrcText) Then s = s & "		<li>源 文 本：" & .SrcText & "</li>" & vbCrLf
				s = s & "	</ul>" & vbCrLf
			End With
			IsErr = True
			'Ebody.Error.Msg = s
			'Ebody.Error.Raise 96
	  End If
	End Function

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

End Class

' xml结点子类
Class Ebody_Xml_Node

	Private o_node

	' 析构
	Private Sub Class_Terminate()
		Set o_node = Nothing
	End Sub

	' 建立新Node对象
	Private Function [New](ByVal o)
		Set [New] = New Ebody_Xml_Node
		[New].Dom = o
	End Function

	' 源对象
	Public Property Let Dom(ByVal o)
		If Not o Is Nothing Then
			Set o_node = o
		Else
			Exit Property
			'Ebody.Error.Msg = "(不是有效的XML对象)"
			'Ebody.Error.Raise 97
		End If
	End Property

	Public Property Get Dom
		Set Dom = o_node
	End Property

	' 取集合中的某一项
	Public Default Function Item(ByVal n)
		' 如果是集合就取其中下标对应子项
		If IsNodes Then
			Set Item = [New](o_node(n))
		' 如果是节点且下标为0就取节点本身
		ElseIf IsNode And n = 0 Then
			Set Item = [New](o_node)
		Else
			Exit Function 
			'Ebody.Error.Msg = "(不是有效的XML元素集合对象&lt;"&TypeName(o_node)&"&gt;)"
			'Ebody.Error.Raise 97
		End If
	End Function

	' =======Xml元素属性（自身属性）======
	' 是否是元素节点
	Public Function IsNode
		IsNode = TypeName(o_node) = "IXMLDOMElement"
	End Function

	' 是否是元素集合
	Public Function IsNodes
		IsNodes = TypeName(o_node) = "IXMLDOMSelection"
	End Function

	' 属性设置(可读可写)
	Public Property Let Attr(ByVal s, ByVal v)
		' 如果值为 Null 相当于删除属性
		If IsNull(v) Then RemoveAttr s : Exit Property
		
		' 判断是否是节点
		If IsNode Then
			' 如果是节点
			o_node.setAttribute s, v		
		ElseIf IsNodes Then
			' 如果是集合则设置每个子节点的属性
			Dim i
			For i = 0 To Length - 1
				o_node(i).setAttribute s, v
			Next
		End If
	End Property

	' 取节点属性
	Public Property Get Attr(ByVal s)
		If Not IsNode Then Exit Property
		Attr = o_node.getAttribute(s)
	End Property

	' 文本属性设置
	Public Property Let Text(ByVal v)
		If IsNode Then
			' 是结点
			If Ebody.Has(v) Then o_node.Text = v		
		ElseIf IsNodes Then
			' 如果是集合则设置每个子节点的文本
			Dim i
			For i = 0 To Length - 1
				If Ebody.Has(v) Then o_node(i).Text = v
			Next
		End If
	End Property

	' 取文本属性
	Public Property Get Text
		If IsNode Then
			Text = o_node.Text
		ElseIf IsNodes Then
			Dim i
			For i = 0 To Length - 1
				Text = Text & o_node(i).Text
			Next
		End If
	End Property

	' 文本内容设置
	Public Property Let Value(ByVal v)
		If IsNode Then
			o_node.ChildNodes(0).NodeValue = v
		ElseIf IsNodes Then
			Dim i
			For i = 0 To Length - 1
				o_node(i).ChildNodes(0).NodeValue = v
			Next
		End If
	End Property

	' 取文本内容设置
	Public Property Get Value
		If Not IsNode Then Exit Property
		Value = o_node.ChildNodes(0).NodeValue
	End Property

	' 获取XML(只读)
	Public Property Get Xml
		If IsNode Then
			Xml = o_node.Xml
		ElseIf IsNodes Then
			Dim i
			For i = 0 To Length - 1
				If i>0 Then Xml = Xml & vbCrLf
				Xml = Xml & o_node(i).Xml
			Next
		End If
	End Property

	' 元素名
	Public Property Get Name
		If Not IsNode Then Exit Property
		Name = o_node.BaseName
	End Property

	' 元素类型
	Public Property Get [Type]
		If IsNodes Then
			[Type] = 0
		Else
			[Type] = o_node.NodeType
		End If
	End Property

	' 元素类型名称
	Public Property Get TypeString
		If IsNodes Then
			TypeString = "selection"
		Else
			TypeString = o_node.NodeTypeString
		End If
	End Property

	' 元素长度
	Public Property Get Length
		If IsNode Then 
			Length = o_node.ChildNodes.Length
		Else
			Length = o_node.Length
		End If
	End Property

	' =======Xml元素属性（返回新节点元素）======
	' 根元素
	Public Function Root
		If IsNode Then
			Set Root = [New](o_node.OwnerDocument)
		Else
			Set Root = [New](o_node(0).OwnerDocument)
		End If
	End Function

	' 父元素
	Public Function Parent
		If Not IsNode Then Exit Function
		Set Parent = [New](o_node.parentNode)
	End Function

	' 子元素
	Public Function Child(ByVal n)
		If Not IsNode Then Exit Function
		Set Child = [New](o_node.ChildNodes(n))
	End Function

	' 上一同级元素
	Public Function Prev
		If Not IsNode Then Exit Function
		Dim o
		Set o = o_node.PreviousSibling
		Do While True
			If TypeName(o) = "Nothing" Or TypeName(o) = "IXMLDOMElement" Then Exit Do
			Set o = o.PreviousSibling
		Loop
		If TypeName(o) = "IXMLDOMElement" Then
			Set [Prev] = [New](o)
			Set o = Nothing
		Else
			Exit Function 
			'Ebody.Error.Msg = "(没有上一同级元素)"
			'Ebody.Error.Raise 96
		End If
	End Function

	' 下一同级元素
	Public Function [Next]
		If Not IsNode Then Exit Function
		Dim o
		Set o = o_node.NextSibling
		Do While True
			If TypeName(o) = "Nothing" Or TypeName(o) = "IXMLDOMElement" Then Exit Do
			Set o = o.NextSibling
		Loop
		If TypeName(o) = "IXMLDOMElement" Then
			Set [Next] = [New](o)
			Set o = Nothing
		Else
			Exit Function 
			'Ebody.Error.Msg = "(没有下一同级元素)"
			'Ebody.Error.Raise 96
		End If
	End Function

	' 第一个元素
	Public Function First
		If Not IsNode Then Exit Function
		Set First = [New](o_node.FirstChild)
	End Function

	' 最后一个元素
	Public Function Last
		If Not IsNode Then Exit Function
		Set Last = [New](o_node.LastChild)
	End Function

	' =======Xml元素方法======
	' (查找)
	' 是否有某属性
	Public Function HasAttr(ByVal s)
		If Not IsNode Then HasAttr = False : Exit Function
		Dim oattr
		Set oattr = o_node.Attributes.GetNamedItem(s)
		HasAttr = Not oattr Is Nothing
		Set oattr = Nothing
	End Function

	' 是否有子节点
	Public Function HasChild()
		If Not IsNode Then HasChild = False : Exit Function
		HasChild = o_node.hasChildNodes()
	End Function

	' 查找子元素
	Public Function Find(ByVal t)
		If Not IsNode Then Exit Function
		Dim o
		If Ebody.Test(t, "[, >\[@:]") Then
			'按简单表达式取元素
			Set o = o_node.selectNodes(Ebody_Xml_TransToXpath(t))
		Else
			'从标签取元素
			Set o = o_node.GetElementsByTagName(t)
		End If
		If o.Length = 0 Then
			Exit Function
			'Ebody.Error.Msg = "("&t&")"
			'Ebody.Error.Raise 98
		ElseIf o.Length = 1 Then
			Set Find = [New](o(0))
		Else
			Set Find = [New](o)
		End If
	End Function

	' XPath取对象集合
	Public Function [Select](ByVal p)
		If Not IsNode Then Exit Function
		Set [Select] = [New](o_node.selectNodes(p))
	End Function

	' XPath取单个对象
	Public Function Sel(ByVal p)
		If Not IsNode Then Exit Function
		Set Sel = [New](o_node.selectSingleNode(p))
	End Function
	
	' (建立)
	' 克隆节点
	Public Function Clone(ByVal b)
		If Not IsNode Then Exit Function
		If Ebody.IsNull(b) Then b = True
		Set Clone = [New](o_node.CloneNode(b))
	End Function

	' 统一对象为Dom节点
	Private Function GetNodeDom(ByVal o)
		Select Case TypeName(o)
			Case "IXMLDOMElement" Set GetNodeDom = o
			Case "Ebody_Xml_Node" Set GetNodeDom = o.Dom
		End Select
	End Function

	' 添加子节点
	Public Function Append(ByVal o)
		If Not IsNode Then Exit Function
		o_node.AppendChild(GetNodeDom(o))
		Set Append = [New](o_node)
	End Function

	' 替换节点
	Public Function ReplaceWith(ByVal o)
		If IsNode Then
			' 如果是节点则直接替换（是Dom内节点会直接移动），返回被替换的节点
			Call o_node.ParentNode.replaceChild(GetNodeDom(o), o_node)
		ElseIf IsNodes Then
			' 如果是集合则依次替换，是Dom内的节点不会移动而是复制
			Dim i,n
			For i = 0 To Length - 1
				Set n = GetNodeDom(o).CloneNode(True)
				Call o_node(i).ParentNode.replaceChild(n, o_node(i))
			Next
		End If
		Set ReplaceWith = [New](o_node)
	End Function

	' 在节点前加入另一个节点
	Public Function Before(ByVal o)
		If IsNode Then
			Call o_node.ParentNode.InsertBefore(GetNodeDom(o), o_node)
		ElseIf IsNodes Then
			Dim i,n
			For i = 0 To Length - 1
				Set n = GetNodeDom(o).CloneNode(True)
				Call o_node(i).ParentNode.InsertBefore(n, o_node(i))
			Next
		End If
		Set Before = [New](o_node)
	End Function

	' 在节点后加入另一个节点
	Public Function After(ByVal o)
		If IsNode Then
			Call InsertAfter(GetNodeDom(o), o_node)
		ElseIf IsNodes Then
			Dim i,n
			For i = 0 To Length - 1
				Set n = GetNodeDom(o).CloneNode(True)
				Call InsertAfter(n, o_node(i))
			Next
		End If
		Set After = [New](o_node)
	End Function

	Private Sub InsertAfter(ByVal n, Byval o)
		Dim p
		Set p = o.ParentNode
		If p.LastChild Is o Then
			p.AppendChild(n)
		Else
			Call p.InsertBefore(n, o.nextSibling)
		End If
	End Sub

	' (删除)
	' 删除某属性
	Public Function RemoveAttr(ByVal s)
		If IsNode Then
			o_node.removeAttribute(s)
		ElseIf IsNodes Then
			Dim i
			For i = 0 To Length - 1
				o_node(i).removeAttribute(s)
			Next
		End If
		Set RemoveAttr = [New](o_node)
	End Function

	' 清空所有子节点
	Public Function [Empty]
		If IsNode Then
			o_node.Text = ""
		ElseIf IsNodes Then
			Dim i
			For i = 0 To Length - 1
				o_node(i).Text = ""
			Next
		End If
		Set [Empty] = [New](o_node)
	End Function

	' 清除所有子节点，包括空文本节点
	Public Function Clear
		If IsNode Then
			o_node.Text = ""
			o_node.removeChild(o_node.FirstChild)
		ElseIf IsNodes Then
			Dim i
			For i = 0 To Length - 1
				o_node(i).Text = ""
				o_node(i).removeChild(o_node(i).FirstChild)
			Next
		End If
		Set Clear = [New](o_node)
	End Function

	' 合并相邻的Text节点并删除空的Text节点
	Public Function Normalize
		If IsNode Then
			o_node.normalize()
		ElseIf IsNodes Then
			Dim i
			For i = 0 To Length - 1
				o_node(i).normalize()
			Next
		End If
		Set Normalize = [New](o_node)
	End Function

	' 删除自身
	Public Sub Remove
		If IsNode Then
			o_node.ParentNode.RemoveChild(o_node)
		ElseIf IsNodes Then
			Dim i
			For i = 0 To Length - 1
				o_node(i).ParentNode.RemoveChild(o_node(i))
			Next
		End If
	End Sub

End Class

Public Function Ebody_Xml_TransToXpath(ByVal s)
	s = Ebody.RegReplace(s, "\s*,\s*", "|//")
	s = Ebody.RegReplace(s, "\s*>\s*", "/")
	s = Ebody.RegReplace(s, "\s+", "//")
	s = Ebody.RegReplace(s, "(\[)([a-zA-Z]+\])", "$1@$2")
	s = Ebody.RegReplace(s, "(\[)([a-zA-Z]+[!]?=[^\]]+\])", "$1@$2")
	s = Ebody.RegReplace(s, "(?!\[\d)\]\[", " and ")
	s = Replace(s, "|", " | ")
	Ebody_Xml_TransToXpath = "//" & s
End Function

%>