<%
'################################################################################
'## ebody.db.asp
'## -----------------------------------------------------------------------------
'## 功能:	数据操纵类
'## 版本:	1.1 Build 130324
'## 作者:	Tony
'## 日期:	2013/03/24
'## 说明:	Ebody基类
'################################################################################
' 所有连接参数都有: 数据源标识,数据库名(路径),IP,端口,用户名,密码
'################################################################################

Class Ebody_db
	
'================================================================================
'== Variable
'================================================================================


'--------------------------------------------------------------------------------
'-- Public
'--------------------------------------------------------------------------------


'--------------------------------------------------------------------------------
'-- Private
'--------------------------------------------------------------------------------

	Private coRS				' RS对像
	Private coConn				' DB对像
	Private cvConnStr			' DB连接字符串
	Private cvSQL				' SQL语句
	Private cvPageSize			' PageSize
	Private cvPageIndex			' 当前页面索引号
	Private cvPageCount			' 分页总数

	Private cvKeyField			' 前台关键栏

'================================================================================
'== Event
'================================================================================

	Private Sub Class_Initialize()
		'On Error Resume Next

		' 默认值
		cvPageSize = 100
		cvPageIndex = 1
		cvPageCount = 1

		' 取默认连接参数
		cvConnStr = Ebody.DBConnStr

	End Sub

	Private Sub Class_Terminate()
		'On Error Resume Next
		CloseRS_()
		CloseDB_()
	End Sub

'================================================================================
'== Public
'================================================================================


'--------------------------------------------------------------------------------
'-- Let Property
'--------------------------------------------------------------------------------
	
	' 设定页面大小
	' 调用: EBody.PageSize = 分页大小
	Public Property Let PageSize(ByVal pNum)
		If IsNumeric(pNum) Then
			cvPageSize = pNum
		End If
	End Property

	' 设定当前页面索引
	' 调用: EBody.PageIndex = 页码
	Public Property Let PageIndex(ByVal pNum)
		If IsNumeric(pNum) Then
			cvPageIndex = Int(pNum)
		End If
	End Property

	' 设定操作基表的唯一关键栏位
	' 调用: EBody.KeyField = 关键栏
	' 说明: 关键栏是行的唯一标识, 所以只能有一个关键栏. 新增, 删除, 修改某些选定的行, 必须指定关键栏
	Public Property Let KeyField(ByVal pStr)
		cvKeyField = pStr
	End Property

	' 设定CONN数据库连接字串
	' 调用: Ebody.ConnStr = 连接字串
	Public Property Let ConnStr(ByVal pStr)
		cvConnStr = pStr
	End Property
	
	

'--------------------------------------------------------------------------------
'-- Get Property
'--------------------------------------------------------------------------------

	' 取得最后一次执行查询后的RS
	Public Property Get RS()
		If IsObject(coRS) Then
			Set RS = coRS
		Else
			RS = coRS
		End If
	End Property

	' 取得数据连接对像
	Public Property Get CONN()
		If IsObject(coConn) Then
			Set CONN = coConn
		Else
			CONN = coConn
		End If
	End Property

	' 取得数据连接字符串
	Public Property Get CONNStr()
		CONNStr = cvConnStr
	End Property

	' 取得执行过的最后一条SQL语句
	Public Property Get SQL()
		SQL = cvSQL
	End Property

	' 取分页总数
	Public Property Get PageCount()
		PageCount = cvPageCount
	End Property

'--------------------------------------------------------------------------------
'-- Function
'--------------------------------------------------------------------------------

	'-------------------------
	' 基础功能类函数
	'-------------------------

	' 创建新对像操作类实例
	' 说明: 当需要创建一个新的实例时可使用此方法
	' 样例: Set loDB = ebody.db.New
	Public Function [New]()
		Set [New] = New ebody_db
	End Function

	' 执行SQL语句
	' 调用 : EBody.ExecuteSQL(SQL语句)
	' 返回 : 查询语句时,返回RS数据集
	' 说明 : 此方法将自动创建并打开DB对像
	Public Function ExecuteSQL(ByVal pSQL)
		'On Error Resume Next

		' 没有SQL,直接退出
		If IsEmpty(pSQL) Then Exit Function

		' 将SQL中形参转为实参(支持SQL中以:打头的形参),只有在非上传Form时才执行
		If Not IsUploadForm_ Then pSQL = GetSqlWithValuesByTag_(pSQL, ":", 0, 0)

		' 打开DB连接
		'Call OpenDB_

		' 开始进程
		coConn.BeginTrans

		' 执行SQL
		If  UCase(GetSqlType_(pSQL))="SELECT" Then

			' 查询操作
			Set coRS = Server.CreateObject("Adodb.Recordset")
			With coRS			
				.ActiveConnection = coConn
				.CursorLocation = 3	' 设定记录集指针属性(oracle分页, 必须为3)
				.CursorType = 1
				.LockType = 1
				.Source = pSQL
				.Open
			End With

			' 记录当前执行SQL
			cvSQL = pSQL

			' 设定数据分页
			Call SetPage_

			' 返回数据对像
			Set ExecuteSQL = coRS

		Else

			' 更新操作
			Dim loCMD : Set loCMD = Server.CreateObject("Adodb.Command")
			With loCMD
				.ActiveConnection = coConn
				.CommandType = 1	' 1 SQL, 4 Procedure
				.CommandText = pSQL
				Set ExecuteSQL = .Execute
			End With
			Set loCMD = Nothing

		End If

		' 判断是否异常, 执行或撤销进程
		If Ebody.IsErr Then
			' 撤销进程
			coConn.RollbackTrans
			Set ExecuteSQL = Nothing
		Else
			' 提交进程
			coConn.CommitTrans			
		End If

		' 关闭DB连接
		'Call CloseDB_
	End Function

	' 根据记录集生成Json格式代码
	' 调用: Ebody.Json(json名称, RS数据对像)
	' 返回: Json字符串内容
	Public Function Json(ByVal pName, ByVal pRS)
		'On Error Resume Next

		'-----------------------------------
		' 使用ebody json类进行转换
		'-----------------------------------
'		Dim loRs, loJson, lvField		
'		' 打开json对像
'		Ebody.Use "json"
'		' 复制出一个数据源副本用于生成json
'		Set loRs = pRS.Clone
'		' 生成一个新json实例, 指定参数为0, 表示是json的对像类型
'		Set loJson = Ebody.Json.New(0)
'		' 将数据拆分后分别存入json数据字典中
'		If Ebody.Has(loRs) Then
'			' 生成新json实例, 指定参数为1, 表示是json的数组类型
'			loJson(pName) = Ebody.Json.New(1)
'			' 取各行的值, 将数据值转存到json字典中, 以便后用
'			While Not loRs.Eof
'				' 生成新json实例, 指定参数为0, 表示是json的对像类型
'				loJson(pName)(Null) = Ebody.Json.New(0)
'				' 取各栏位的值
'				For Each lvField In loRs.Fields					
'					' 将栏位值存入json对像
'					loJson(pName)(Null)(lvField.Name) = lvField.Value
'				Next
'				' 下一笔
'				loRs.MoveNext
'			Wend
'		End If
'		' 生成json数据, 并返回
'		Json = loJson.JsString
'
'		' 释放对像
'		Set loJson = Nothing
'		loRs.Close() : Set loRs = Nothing


		'-----------------------------------
		' 以下只针对数据集简易生成json功能，执行速度更快
		'-----------------------------------

		If LCase(TypeName(pRS)) = "recordset" Then
		
			Dim lvTotal, lvField, lvJson, lvJsonBody, lvFlag1, lvFlag2, lvCount

			lvTotal = pRS.RecordCount
			lvFlag2 = True

			' 组构所有记录为json字串			
			While (Not pRS.Eof) And (lvCount < pRS.PageSize)

				' 初始化
				lvFlag1 = True
				lvJsonBody = Empty

				' 组构所有栏位
				For Each lvField In pRS.Fields
					If lvFlag1 Then lvFlag1 = False Else lvJsonBody = lvJsonBody & ","
					'lvJsonBody = lvJsonBody & """" & lvField.Name & """" & ":" & """" & lvField.Value & """"	' disabled by tony 20140326
					lvJsonBody = lvJsonBody & lvField.Name & ":" & ToJSON_(lvField.Value)	' add by tony 20140326
				Next

				If lvFlag2 Then lvFlag2 = False Else lvJson = lvJson & ","
				lvJson = lvJson & "{" & lvJsonBody & "}"	
				lvCount = lvCount + 1
				pRS.MoveNext

			Wend

		End If

		' 返回json字串
		'Json = "{""" & pName & """:[" & lvJson & "]}"	' disabled by tony 20140326
		Json = pName & ":[" & lvJson & "]"		' add by tony 20140326
		
	End Function


	'-------------------------
	' 验证类函数(Is打头)
	'-------------------------

	' 判断对像是否已经打开
	' 调用: IsOpen(对像)
	' 返回: true 已打开, false 未打开
	Public Function IsOpen(ByRef pObj)
		Select Case TypeName(pObj)
		Case "Connection","Recordset"
			If pObj.State = 1 Then IsOpen = True Else IsOpen = False
		Case Else
			IsOpen = False
		End Select
	End Function

	'-------------------------
	' 系统取值类函数(Get打头)
	'-------------------------

	' 取得指定行列坐标的值
	' 调用: EBody.GetValue(行号,列号)
	' 返回: 绝对位行列的值
	' 说明: 性能低,只适用于少量数据
	Public Function GetValue(ByVal pRowNum, ByVal pColNum)
		If coRS.Eof Or coRS.Bof Then Exit Function

		' 记录原定位行号
		Dim lvPreRowNum : lvPreRowNum = coRS.absoluteposition		

		' 将记录指针移到数据表第N行
		coRS.AbsolutePosition = pRowNum

		' 另一种定位方法,性能低,只适用于少量数据
		'Dim lvRowNum : lvRowNum = 0
		' 初始定位到首行
		'coRS.MoveFirst
		' 定位指定行列
		'Do While Not coRS.Eof Or coRS.Bof			
		'	lvRowNum = lvRowNum + 1
		'	If lvRowNum = pRowNum Then Exit Do
		'	coRS.MoveNext
		'Loop

		' 返回值(取某行某例的值)
		GetValue = coRS.Fields(pColNum-1).value

		' 指针移到原定位行
		coRS.AbsolutePosition = lvPreRowNum
	End Function

	' 取得指定列名称
	' 调用: EBody.GetColName(列号)
	' 返回: 指定列的列名
	Public Function GetColName(ByVal pColNum)
		Dim lvCount, lvColName
		' 依参数定位列
		For lvCount = 0 To coRS.Fields.Count - 1
			If lvCount = pColNum - 1 Then 
				lvColName = coRS.Fields(lvCount).name
				Exit For
			End If
		Next
		' 返回指定的列名
		GetColName = lvColName
	End Function


'--------------------------------------------------------------------------------
'-- Sub
'--------------------------------------------------------------------------------
	
	'-------------------------
	' 基础功能类过程
	'-------------------------

	' 设置DB连接参数
	' 说明: 类同于ConnStr属性方法,只不过这种方式是格式化的设定方法
	' 调用: SetConn(数据库类型, 数据源标识, 数据库资源(路径), 服务器地址(IP), 端口, 用户, 密码)
	Public Sub SetConn(ByVal pType, ByVal pID, ByVal pSource, ByVal pHost, ByVal pPort, ByVal pUser, ByVal pPass)
		Select Case UCase(pType)
			Case "0","MSSQL"
				cvConnStr = "PROVIDER=SQLOLEDB;DATA SOURCE=" & pHost & ";UID=" & pUser & ";PWD=" & pPass & ";DATABASE=" & pID & ";"
			Case "1","ACCESS"
				cvConnStr = "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=" & pSource & ";JET OLEDB:DATABASE PASSWORD=" & pPass & ";"
			Case "2","MYSQL"
				cvConnStr = "DRIVER={MYSQL ODBC 3.51 DRIVER};SERVER=" & pHost & ";PORT=" & pPort & ";DATABASE=" & pSource & ";USER NAME=" & pUser & ";PASSWORD=" & pPass &";"
			Case "3","ORACLE"
				cvConnStr = "PROVIDER=MSDAORA;DATA SOURCE=(DESCRIPTION =(ADDRESS_LIST =(ADDRESS = (COMMUNITY = TCP.WORLD)(PROTOCOL = TCP)(HOST = " & pHost & ")(PORT = " & pPort & ")))(CONNECT_DATA =(SID = " & pID & "))); USER ID=" & pUser & "; PASSWORD=" & pPass & ";"
			Case "4","EXCEL"
				cvConnStr = "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=" & pSource & ";EXTENDED PROPERTIES=EXCEL 8.0;"
		End Select
	End Sub

	' 打开数据库
	' 调用: Ebody.Open
	' 说明: 执行前需要先设制conn连接字串
	Public Sub Open()
		' 打开DB连接
		OpenDB_
	End Sub

	' 关闭数据库连接对像, 释放内存
	Public Sub Close()
		CloseDB_()
		CloseRS_()
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

	' 将数据转化成适应于Json字符串的格式
	' 说明: 一般字符串类型的要加引号，而数值类型不需加引号
	Private Function ToJSON_(pData)
		Select Case VarType(pData)
			Case 1 ' Null（无有效数据）
				toJSON_ = "null"
			Case 7 ' 日期
				toJSON_ = """" & CStr(pData) & """"
			Case 8 ' 字符串
				toJSON_ = """" & Ebody.Escape(pData) & """"
			Case 11 ' Boolean
				toJSON_ = Ebody.IIF(pData, "true", "false")
			Case Else
				toJSON_ = Replace(pData, ",", ".")
		End select
	End Function

	'-------------------------
	' 验证类函数(Is打头)
	'-------------------------

	' 检查表单上传类型是否为:multipart/form-data
	Private Function IsUploadForm_()
		Dim lvRequestMethod
		lvRequestMethod = trim(LCase(Request.ServerVariables("REQUEST_METHOD")))
		If lvRequestMethod="" or lvRequestMethod<>"post" Then
			IsUploadForm_ = False
			Exit Function
		End If
		Dim FormType : FormType = Split(Request.ServerVariables("HTTP_CONTENT_TYPE"), ";")
		If LCase(FormType(0))<>"multipart/form-data" Then
			IsUploadForm_ = False
		Else
			IsUploadForm_ = True
		End If
	End Function

	'-------------------------
	' 系统取值类函数(Get打头)
	'-------------------------

	' 返回带实参的SQL语句,以实参替换SQL中带标记符的形参(此功能仿参数化查询)
	' 调用:	GetSqlWithValuesByTag_(SQL语句, 标记字符[用以标记其为参数栏位], 行号[0表示所有行,其它数值则表示对应数值的行], 是否自动匹配格式)
	' 样例:	GetSqlWithValuesByTag_("select * from t_user where user_id in (:userid)",":",0,1)
	' 返回:	参数栏位替换为其前台对应栏位值的新SQL语句
	' 参数:	pSQL 将要被替换的语句
	'		pTag 形参标识符, 是以传来的以此参数字符开头的栏位名, 对应的值是提取前台由Post Form提交过来的值
	'		pRowNum 行号, 0=全部行值, 其它值=指定的行值, 如果行号为0, 并且前台值有多个, 则实参值就是以, 
	'				号分隔的的字串(相当于SQL中in里的值清单). 如行号为指定行, 则实参值为前台提交指定行的栏位值
	'		pAutoFormat 是否自动匹配格式, 1=自动匹配, 0=不自动匹配, 自动匹配会自动为字符串值加'(单引号), 一般用于where条件中的in值集合; 
	'					不自动区配, 则需要在属于字符串的形参前后加'(单引号)来说明他是字符串
	' 说明: 此方法可防SQL注入
	Private Function GetSqlWithValuesByTag_(ByVal pSQL, ByVal pTag, ByVal pRowNum, ByVal pAutoFormat)
		Dim lvFields, lvNewSQL, lvField, lvValues, lvValue, lvValueSet, lvValudation, lvSplitStr
		' 取得SQL中的所有参数栏
		lvFields = Split(GetFieldsByTag_(pSQL,pTag), ",")
		lvNewSQL = pSQL
		' 将SQL中各栏位名替换为前台提交来的实值
		For Each lvField In lvFields
			' 取得提交来的参数值
			If IsEmpty(Request.Form(lvField)) Then
				' 通过地址传参, 则直接取地址栏中的参数值
				ReDim lvValues(0) : lvValues(0) = Request(lvField)
			Else			
				' 通过Form传参取参, 取Form中的所有参数值
				' 注意: 只有关键字段才取值集, 以供In条件使用
				If pRowNum=0 And lvField=cvKeyField Then
					' 取得某列的值集(如:a,b,c,d)
					' 组构成组
					lvValues = Split(Request.Form(lvField), ",")
				Else
					' 取Form中指定行数据的某列值(绝对定位取值)
					If pRowNum=0 Then pRowNum = 1	' 如果取form中的值, 但又指定的是0, 则改为1, 默认取第一个值
					ReDim lvValues(0) : lvValues(0)= Request.Form(lvField)(pRowNum)
				End If
			End If
			' 组构值集
			' 如果行号为0, 则提交的参数是一个值集, 以使用于in条件中.
			' 此时在值集中, 需判断值集的类型, 并将值集中的值做相应的转换
			lvValueSet = empty	' 值集初始(必须)
			lvSplitStr = empty	' 分隔符初始(必须)
			For Each lvValue In lvValues
				' 1. 替换字符串中的单引号
				' 防注入SQL代码(有效防注入90%以上)
				lvValue = Replace(Trim(lvValue), "'", "''")

				' 2. 组构值集, 自动匹配时, 验证值的类型, 字符型则转为字符型(即在值两边加上'号)
				If IsNumeric(lvValue) Or IsDate(lvValue) Or pAutoFormat=0 Then
					' 另一种写法
					lvValueSet = lvValueSet & lvSplitStr & lvValue
				Else 
					lvValueSet = lvValueSet & lvSplitStr & "'" & lvValue & "'"
				End If
				lvSplitStr = ","
			Next
			' 替换SQL中的各参数为值
			' pTag&"+\b"&lvField&"\b"是正则表达式规则, 即在SQL中寻找带pTag为开头的栏位
			lvNewSQL = RegReplace(lvNewSQL, pTag&"+\b"&lvField&"\b", lvValueSet)
		Next		
		' 返回
		GetSqlWithValuesByTag_ = lvNewSQL
	End Function

	' 取SQL中的各个形参名称
	' 在SQL字串中取得以指定以标识字符开头的前台栏位集字串(仿正则表达式取值)
	' 调用: GetFieldsByTag_(要分析的SQL字符串, 分隔符)
	' 样列: GetFiledsByTag("select * from t_user where user_name=:user_name and user_id=:user_id",":")
	'		返回: user_name,user_id
	' 返回: 形参栏位集字串
	' 说明: 形参栏是指前面带有指定标记的栏位.一般应用于SQL语句中取形参名称.在SQL中,带:号的字段名表示前台提交的栏位
	Private Function GetFieldsByTag_(ByVal pSQL, ByVal pTag)
		Dim lvBeginPos, lvEndPos, lvEndTags, lvEndTag, lvFields, lvField, lvEndPosPre, lvEndPosNew, lvEndPosOld, lvSplitStr
		Dim loDicField : Set loDicField = Server.CreateObject("Scripting.Dictionary") : loDicField.CompareMode = 1
		' 结束标识符集
		lvEndTags = Split(Chr(Asc("%"))&"|"&Chr(Asc("'"))&"|"&Space(1)&"|,|)|;|*|/|+|-|<|>","|")
		lvBeginPos = 1
		Do
			' 取得标记开始位,初始值
			lvBeginPos = InStr(lvBeginPos, pSQL, pTag)
			lvEndPos = 0
			lvEndPosNew = 0
			lvEndPosOld = 0
			If lvBeginPos > 0 Then				
				' 取离查询栏位最近的一个结束标记位
				For Each lvEndTag In lvEndTags
					If lvEndPosNew > 0 Then lvEndPosOld = lvEndPos	' 记录前一个结束位
					lvEndPosNew = InStr(lvBeginPos,pSQL,lvEndTag)	' 取最新结束位
					If lvEndPosNew = 0 Then lvEndPosNew = lvEndPosOld	' 如新结束位为0,则以旧值替代
					If lvEndPosOld > 0 And lvEndPosOld < lvEndPosNew Then lvEndPos = lvEndPosOld Else lvEndPos = lvEndPosNew	' 比较新旧值,存入小的结束位
				Next				
				' 取SQL语句中带:号的栏位名(形参,即前台栏位名)
				If lvEndPos > 0 Then
					lvField = Mid(pSQL,lvBeginPos+1,lvEndPos-1-lvBeginPos)
				Else
					lvField = Mid(pSQL,lvBeginPos+1)
				End If
				' 组构形参栏位
				' 通过数据字典对像验证是否重复的形参栏,重复栏只取一次
				If Not loDicField.Exists(lvField) Then loDicField.add lvField, lvField : lvFields = lvFields & lvSplitStr & lvField
				lvSplitStr = ","
				' 重新得到下个起点位
				lvBeginPos = lvEndPos				
			End If
		Loop While lvBeginPos > 0
		' 释放字典
		loDicField.RemoveAll
		Set loDicField = Nothing
		' 返回
		GetFieldsByTag_ = lvFields
	End Function

	' 取得SQL的类型
	' 调用: GetSqlType_(SQL语句)
	' 返回: select, insert, update, delete, grant, drop, unknow
	' 说明: 依SQL语句取得其操作类型
	Private Function GetSqlType_(ByVal pSQL)
		If Len(pSQL) = 0 Then Exit Function
		pSQL = Trim(pSQL)
		' 去掉Tag符
		GetSqlType_ = Replace(LCase(Left(pSQL, InStr(pSQL, " ") - 1)), Chr(9), "")
	End Function

	' 取得数据库类型
	' 调用: GetDBType_(连接对像)
	' 返回: 数据库类型字符串
	Private Function GetDBType_(ByVal pConn)
		Dim lvProvider : lvProvider = UCase(pConn.Provider)
		Dim lvCount, MSSQL, ACCESS, MYSQL, ORACLE
		MSSQL = Split("SQLNCLI10, SQLXMLOLEDB, SQLNCLI, SQLOLEDB, MSDASQL",", ")
		ACCESS = Split("MICROSOFT.ACE.OLEDB.12.0, MICROSOFT.JET.OLEDB.4.0",", ")
		MYSQL = "MYSQLPROV"
		ORACLE = Split("MSDAORA, OLEDB.ORACLE",", ")
		For lvCount = 0 To Ubound(MSSQL)
			If Instr(lvProvider,MSSQL(lvCount))>0 Then
				GetDBType_ = "MSSQL" : Exit Function
			End If
		Next
		For lvCount = 0 To Ubound(ACCESS)
			If Instr(lvProvider,ACCESS(lvCount))>0 Then
				GetDBType_ = "ACCESS" : Exit Function
			End If
		Next
		If Instr(lvProvider,MYSQL)>0 Then
			GetDBType_ = "MYSQL" : Exit Function
		End If
		For lvCount = 0 To Ubound(ORACLE)
			If Instr(lvProvider,ORACLE(lvCount))>0 Then
				GetDBType_ = "ORACLE" : Exit Function
			End If
		Next


	End Function

'--------------------------------------------------------------------------------
'-- Sub
'--------------------------------------------------------------------------------

	'-------------------------
	' 基础功能类过程
	'-------------------------

	' 打开数据库连接对像
	Private Sub OpenDB_()
		'On Error Resume Next
		If TypeName(coConn) = "Connection" Then Exit Sub
		Set coConn = Server.Createobject("Adodb.Connection")
		coConn.Open cvConnStr
		If Ebody.IsErr Then
			coConn.Close
			Set coConn = Nothing
		End If
	End Sub

	' 关闭数据库连接对像
	Private Sub CloseDB_()
		'On Error Resume Next
		If TypeName(coConn) = "Connection" Then
			'If coConn.State = 1 Then coConn.Close
			Set coConn = Nothing
			Err.Clear
		End If		
	End Sub
	
	' 关闭数据集对像
	Private Sub CloseRS_()
		'On Error Resume Next
		If TypeName(coRS) = "Recordset" Then
			'If coRS.State = 1 Then coRS.Close
			Set coRS = Nothing
			coRS = Empty
			Err.Clear
		End If
	End Sub

	'-------------------------
	' 设置对像类过程(Set打头)
	'-------------------------

	' 设定分页参数
	' 说明: 设定数据分页及定位当前索引页面
	Private Sub SetPage_()
		' 验证
		If UCase(GetSqlType_(cvSQL))<>"SELECT" Then Exit Sub
		If Not IsOpen(coRS) Then Exit Sub		

		' 设定数据分页参数
		Select Case GetDBType_(coConn)
		Case "ACCESS","ORACLE","MSSQL"
			' 设定分页大小
			coRS.PageSize = cvPageSize
			' 算出记录总页数
			cvPageCount = coRS.PageCount
			If cvPageCount<1 Or cvPageIndex>cvPageCount Then
				Exit Sub
			Else
				' 设定当前页面位置
				coRS.AbsolutePage = cvPageIndex
			End If
		Case Else
			Exit Sub
		End Select
	End Sub

End Class
%>