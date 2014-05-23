<%
'################################################################################
'## ebody.json.asp
'## -----------------------------------------------------------------------------
'## 功能:	JSON生成类
'## 版本:	1.1 Build 130324
'## 作者:	Tony
'## 日期:	2013/03/24
'## 说明:	Ebody基类
'################################################################################

Class ebody_json
	
'================================================================================
'== Variable
'================================================================================

	Public QuotedVars		' 是否使用引号
	Public StrEncode		' 中文是否编码
	Public Kind				' JSON数据的类型,object会用{}包含,array则用[]包含
	Public Collection		' 存储JSON项

'--------------------------------------------------------------------------------
'-- Public
'--------------------------------------------------------------------------------
		
	

'--------------------------------------------------------------------------------
'-- Private
'--------------------------------------------------------------------------------

	Private cvCount			' 对像计数
	Private coJson			' Json解析对像

'================================================================================
'== Event
'================================================================================

	Private Sub Class_Initialize()
		'On Error Resume Next

		' 创建存储
		Set Collection = CreateObject("Scripting.Dictionary")

		' 创建json解析对像
		InitScriptControl_

		' 名称是否用引号,中文是否编码
		If TypeName(Ebody.Json) = "ebody_json" Then
			QuotedVars = Ebody.Json.QuotedVars
			StrEncode = Ebody.Json.StrEncode
		Else
			QuotedVars = True
			StrEncode = True
		End If
		
		' 初始化
		cvCount = 0

	End Sub

	Private Sub Class_Terminate()
		'On Error Resume Next
		Set Collection = Nothing

		' 销毁json解析对像
		DestroyScriptControl_
	End Sub

'================================================================================
'== Public
'================================================================================


'--------------------------------------------------------------------------------
'-- Let Property
'--------------------------------------------------------------------------------

	' 设置JSON项的值(默认调用的属性)
	Public Property Let Pair(ByRef pName, ByRef pData)

		' 如果名称为空,则使用默认的值
		If IsNull(pName) Then pName = cvCount : cvCount = cvCount + 1 ' 预算出下一个计数
		
		' 将数据写入字典
		If vartype(pData) = 9 Then
			' vbObject对像
			If TypeName(pData) = "ebody_json" Then
				Set Collection(pName) = pData

			Else
				Collection(pName) = pData
			End If
		Else
			' 其它对像
			Collection(pName) = pData
		End If
	End Property

'--------------------------------------------------------------------------------
'-- Get Property
'--------------------------------------------------------------------------------

	' 读取JSON项的值(默认调用的属性)
	Public Default Property Get Pair(ByRef pName)

		' 如果名称为空,则使用默认的值
		If IsNull(pName) Then pName = cvCount - 1 ' 因为在let pair中有预+1,所以要读取当前值,则需要-1

		If IsObject(Collection(pName)) Then
			Set Pair = Collection(pName)
		Else
			Pair = Collection(pName)
		End If

	End Property

	' 返回Json字符串
	Public Property Get JSON
		JSON = toJSON(Me)
	End Property
	
'--------------------------------------------------------------------------------
'-- Function
'--------------------------------------------------------------------------------

	'-------------------------
	' 基础功能类函数
	'-------------------------

	' 建新JSON类实例
	' 说明: 当需要创建一个新的实例时可使用此方法
	' 样例: Set loJson = ebody.json.New
	'Public Function [New]
	'	Set [New] = New ebody_json
	'End Function

	Public Function [New](ByVal pKind)
		Set [New] = New ebody_json
		Select Case LCase(pKind)
			Case "0", "object" [New].Kind = 0
			Case "1", "array"  [New].Kind = 1
		End Select
	End Function

	' 将数据转化成Json字符串
	Public Function ToJSON(pData)
		Select Case VarType(pData)
			Case 0	'未初始化，空字符
				toJSON = """"""
			Case 1 ' Null（无有效数据）
				toJSON = "null"
			Case 2, 3, 4, 5 ' 数值类型
				toJSON = pData
			Case 7 ' 日期
				toJSON = """" & CStr(pData) & """"
			Case 8 ' 字符串
				'pData = Ebody.Escape(pData)
				'StrEncode = false
				toJSON = """" & Ebody.IIF(StrEncode, Ebody.Escape(pData), JsEncode_(pData)) & """"
				'toJSON = Ebody.UnEscape(toJSON)
			Case 9 ' vbObject
				Dim lvIsSplit, i 
				lvIsSplit = True
				toJSON = toJSON & Ebody.IIF(pData.Kind, "[", "{")
				For Each i In pData.Collection
					If lvIsSplit Then lvIsSplit = False Else toJSON = toJSON & ","
					toJSON = toJSON & Ebody.IfThen(pData.Kind=0, Ebody.IIF(QuotedVars, """" & i & """", i) & ":") & toJSON(pData(i))
				Next
				toJSON = toJSON & Ebody.IIF(pData.Kind, "]", "}")
			Case 11 ' Boolean
				toJSON = Ebody.IIF(pData, "true", "false")
			Case 12, 8192, 8204 ' 变量数组
				toJSON = RenderArray_(pData, 1, "")
			Case Else
				toJSON = Replace(pData, ",", ".")
		End select
	End Function

	' 复制Json对象,返回一个新的副本
	Public Function Clone
		Set Clone = ColClone_(Me)
	End Function	

	' 依字串生成json对像
	Public Function parse(ByRef pJsonStr)
		Set parse = getJSONObject_(pJsonStr)
	End Function

	'-------------------------
	' 验证类函数(Is打头)
	'-------------------------

	'-------------------------
	' 系统取值类函数(Get打头)
	'-------------------------

	' 取json数组值
	Public Function getValue(ByRef pObjJSArray, ByRef pIndex)
		getValue = getJSArrayValue_(pObjJSArray, pIndex)
	End Function	

'--------------------------------------------------------------------------------
'-- Sub
'--------------------------------------------------------------------------------
	
	'-------------------------
	' 基础功能类过程
	'-------------------------

	' 清除所有JSON项
	Public Sub Clean
		Collection.RemoveAll
	End Sub

	' 删除某一JSON项值
	Public Sub Remove(ByRef pName)
		Collection.Remove pName
	End Sub

	' 输出为Json格式文件到页面上
	Public Sub Flush
		Response.Clear()
		Response.Charset = "UTF-8"
		Response.ContentType = "application/json"
		Ebody.NoCache()
		Response.write JSON
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

	' 处理字符串中的Javascript特殊字符，不处理中文
	Private Function JsEncode_(ByVal s)
		If Ebody.isNull(s) Then JsEncode_ = "" : Exit Function
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
			If p Then t = t & c
		Next
		JsEncode_ = t
	End Function

	' 递归数组生成Json字符串
	Private Function RenderArray_(arr, depth, parent)
		Dim first : first = LBound(arr, depth)
		Dim last : last = UBound(arr, depth)
		Dim index, rendered
		Dim limiter : limiter = ","
		RenderArray_ = "["
		For index = first To last
			If index = last Then
				limiter = ""
			End If 
			On Error Resume Next
			rendered = RenderArray_(arr, depth + 1, parent & index & "," )
			If Err = 9 Then
				On Error GoTo 0
				RenderArray_ = RenderArray_ & toJSON(Eval("arr(" & parent & index & ")")) & limiter
			Else
				RenderArray_ = RenderArray_ & rendered & "" & limiter
			End If
		Next
		RenderArray_ = RenderArray_ & "]"
	End Function

	' 复制JSON对像内容
	Private Function ColClone_(ByRef pObj)
		Dim loJson, i
		Set loJson = new ebody_json
		loJson.Kind = pObj.Kind
		For Each i In pObj.cvCollection
			If IsObject(pObj(i)) Then
				Set loJson(i) = ColClone_(pObj(i))
			Else
				loJson(i) = pObj(i)
			End If
		Next
		Set ColClone_ = loJson
	End Function

	'-------------------------
	' 验证类函数(Is打头)
	'-------------------------

	'-------------------------
	' 系统取值类函数(Get打头)
	'-------------------------

	' 取得对像计数,将用在数据字典中
	Private Function GetCounter()
		GetCounter = cvCount ' 返回
		cvCount = cvCount + 1 ' 预算出下一个计数
	End Function


	' 将Json字串转换为Json对像
	' 返回: json对像
	' 调用: Set objTest = getJSONObject_(Json字串)
	' 说明: 执行后，可使用对像.的方法取得json中的值，如：objTest.name
	Private Function getJSONObject_(ByRef pJsonStr)
	    coJson.AddCode "var jsonObject = " & pJsonStr
		Set getJSONObject_ = coJson.CodeObject.jsonObject
	End Function


	' 取json对像中数组中的某值 
	' 调用: getJSArrayValue_(json对像, 数组索引)
	' 返回: 指定数组内的值
	Private Function getJSArrayValue_(ByRef pObjJSArray, ByRef pIndex)
		'On Error Resume Next
		Dim loDest
		coJson.Run "getJSArray", pObjJSArray, pIndex
		'Set getJSArrayValue_ = coJson.CodeObject.itemTemp
		'If Err.number=0 Then Exit Function
		getJSArrayValue_ = coJson.CodeObject.itemTemp
	End Function

'--------------------------------------------------------------------------------
'-- Sub
'--------------------------------------------------------------------------------

	'-------------------------
	' 基础功能类过程
	'-------------------------

	' 创建解析Json的对像
	Private Sub InitScriptControl_()
		Set coJson = Server.CreateObject("MSScriptControl.ScriptControl")
			coJson.Language = "JavaScript"
			coJson.AddCode "var itemTemp=null;function getJSArray(arr, index){itemTemp=arr[index];}"
	End Sub

	' 销毁解析Json的对像
	Private Sub DestroyScriptControl_()
		Set coJson = Nothing
	End Sub

	'-------------------------
	' 设置对像类过程(Set打头)
	'-------------------------

End Class
%>