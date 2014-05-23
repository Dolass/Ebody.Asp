<%
'################################################################################
'## ebody.tpl.asp
'## -----------------------------------------------------------------------------
'## 功能:	TPL模板
'## 版本:	1.1 Build 130324
'## 作者:	Tony
'## 日期:	2013/03/24
'## 说明:	Ebody基类
'################################################################################
'## TPL功能及技巧介绍:
'## -----------------------------------------------------------------------------
'## 1. 如果块没有更新(UpdateBlock),虽然前面有加载(或者说是绑定)数据源(LoadData),也不会更新块内的数据.
'## 2. 您可以绑定一个值到标签上,如果这个标签刚好在另一个数据块(RS块)内也存在,那么系统会优先使用先前单独绑定的值来更新数据块内的相同标签.利用这个功能,我们可以成批的替换掉数据块中的任意栏位.(如:可用于某些栏位针对某些人员不便显其内容,而用***代替显示等,起到保密作用.)
'## 3. TPL可以绑定数据(RS)块,这样可以更简单的显示出RS中的每条数据行.
'## 4. 允许取任意一个块内的原始数据,或替换后的数据
'## 注意:
'## -----------------------------------------------------------------------------
'## 1. Tpl模板文件中不能定义两个相同名称的块,否则系统会依最后一个块为准
'################################################################################
Class ebody_tpl
	
'================================================================================
'== Variable
'================================================================================


'--------------------------------------------------------------------------------
'-- Public
'--------------------------------------------------------------------------------


'--------------------------------------------------------------------------------
'-- Private
'--------------------------------------------------------------------------------
	Private cvHtml			' 原模板内容
	Private coTagData		' 标签数据字典
	Private coBlockBody		' 块内容字典
	Private coBlockBodyMark	' 新块内容字典
	Private coBlockData		' 已更新数据块字典
	Private cvUnDefine		' 未定义标签的显示方式
	Private cvFile			' tpl模板文件地址

'================================================================================
'== Event
'================================================================================

	Private Sub Class_Initialize()
		'On Error Resume Next
		' 记录标签对应实值字典
		Set coTagData = Server.CreateObject("Scripting.Dictionary") : coTagData.CompareMode = 1
		' 记录各块内容字典
		Set coBlockBody = Server.CreateObject("Scripting.Dictionary") : coBlockBody.CompareMode = 1
		' 新块内容字典
		Set coBlockBodyMark = Server.CreateObject("Scripting.Dictionary") : coBlockBodyMark.CompareMode = 1
		' 已更新数据块字典
		Set coBlockData = Server.CreateObject("Scripting.Dictionary") : coBlockData.CompareMode = 1
	End Sub

	Private Sub Class_Terminate()
		'On Error Resume Next
		coTagData.RemoveAll
		Set coTagData = Nothing
		coBlockBody.RemoveAll
		Set coBlockBody = Nothing
		coBlockBodyMark.RemoveAll
		Set coBlockBodyMark = Nothing
		coBlockData.RemoveAll
		Set coBlockData = Nothing
	End Sub

'================================================================================
'== Public
'================================================================================


'--------------------------------------------------------------------------------
'-- Let Property
'--------------------------------------------------------------------------------

	' 设定未定义标签的显示方式
	' 调用: UnDefine = 显示方式
	' 样例: Ebody.tpl.UnDefine = "remove"
	Public Property Let UnDefine(ByVal pType)
		Select Case LCase(pType)
		Case "remove"
			cvUnDefine = "remove"
		Case "comment"
			cvUnDefine = "comment"
		Case "keep"
			cvUnDefine = "keep"
		Case Else
			cvUnDefine = "keep"
		End Select
	End Property

	' 通过属性的方式载入标签数据
	' 调用: tag(标签名) = 标签值
	' 样例: Ebody.tpl.tag("title") = "welcome"
	' 说明: 此方法是为了摸拟oo的操作方法
	Public Property Let Tag(ByVal pTagName, ByVal pValue)
		Call LoadData_(pTagName,pValue)
	End Property

	' 加载TPL模板文件
	' 调用: File = 文件地址
	' 样例: Ebody.tpl.File = "c:\tpl\abc.htm"
	Public Property Let File(ByVal pFilePath)
		cvFile = pFilePath
		Call LoadTpl_(pFilePath)
	End Property

	' 加载文本字符串模板
	' 调用: Str = 字符串
	' 样例: Ebody.tpl.Str = "this is demo {filename} for ebody110"
	Public Property Let Str(ByVal pStr)
		Call LoadStr_(pStr)
	End Property

'--------------------------------------------------------------------------------
'-- Get Property
'--------------------------------------------------------------------------------

	' 取得TPL模板文件地址
	' 调用: File
	' 返回: TPL模板文件地址
	' 样例: Ebody.tpl.File
	Public Property Get File()
		File = cvFile
	End Property

	' 取得填充后的HTML内容
	' 调用: Ebody.Tpl.HTML
	Public Property Get HTML()
		HTML = cvHtml
	End Property

	' 取得已填充后标签数据
	' 调用: tag(标签名)
	' 样例: Ebody.tpl.tag("title")
	Public Property Get TagData(ByVal pTagName)
		If coTagData.Exists(pTagName) Then
			TagData = coTagData.Item(pTagName)
		Else
			TagData = ""
		End If
	End Property

	' 取得原始块数据
	' 调用: block(块名)
	' 样例: Ebody.tpl.block("page")
	Public Property Get Block(ByVal pBlockName)
		If coBlockBody.Exists(pBlockName) Then
			Block = coBlockBody.Item(pBlockName)
		Else
			Block = ""
		End If
	End Property

	' 取得已填充后的当前块层数据(不返回子块内容)
	' 调用: BlockData(块名)
	' 样例: Ebody.tpl.BlockData("page")
	Public Property Get BlockData(ByVal pBlockName)
		If coBlockData.Exists(pBlockName) Then
			BlockData = coBlockData.Item(pBlockName)
		Else
			BlockData = ""
		End If
	End Property

	' 取得已填充后的块内所有数据(同时返回子块下的所有填充后的内容)
	' 调用: BlockDataAll(块名)
	' 样例: Ebody.tpl.BlockDataAll("page")
	Public Property Get BlockDataAll(ByVal pBlockName)
		If coBlockData.Exists(pBlockName) Then
			BlockDataAll = GetBlockDataAll_("---" & pBlockName & "---")
		Else
			BlockDataAll = ""
		End If
	End Property

	 
'--------------------------------------------------------------------------------
'-- Function
'--------------------------------------------------------------------------------

	' 创建新对像操作类实例
	' 说明: 当需要创建一个新的实例时可使用此方法
	' 样例: Set loTpl = ebody.tpl.New
	Public Function [New]()
		Set [New] = New ebody_tpl
	End Function

	' 生成填充后的HTML内容
	' 调用: GetHtml()
	' 返回: 完整的HTML内容
	Public Function GetHtml()
		' 1. 绑定所有块中的标签值
		Call UpdateBlocks_
		' 2. 更新标签数据
		Call BuildTags_
		' 3. 更新块数据
		Call BuildBlocks_
		' 返回
		GetHtml = cvHtml
	End Function

'--------------------------------------------------------------------------------
'-- Sub
'--------------------------------------------------------------------------------

	' 绑定块中的标签值
	' 调用: UpdateBlock(块标记名)
	' 说明: 此方法依据绑定的标签值更新各块内容,并存入字典coBlockData
	Public Default Sub UpdateBlock(ByVal pBlockName)
		Dim lvRule, lvBlockName, lvBlockData, lvRowNum, lvFieldNum
		Dim lvTagDataType
		' 更新块内容
		' 此处用于取出当初存储RS对像的字典Item值,然后循环更新各数据行
		' 注意: TypeName(coTagData.Item(pBlockName))这段语句会自动新增一个字典值,所以先判断一下字典对像是否存在,以避免造成可能的错误
		If coTagData.Exists(pBlockName) Then lvTagDataType = TypeName(coTagData.Item(pBlockName))
		If lvTagDataType = "Recordset" Then
			' 更新块内的数据集,在之前的LoadData_方法中,用到coTagData其中的一个item临时存储RS来循环显出各数据行
			Dim loRS : Set loRS = coTagData.Item(pBlockName)
			Dim lvPageSize : lvPageSize = loRS.PageSize
			Dim lvTagName, lvTagNum, lvFieldValue
			ReDim lvHasTags(loRS.Fields.Count - 1)
			' 依据分页大小组构RS块内容
			For lvRowNum=1 To lvPageSize
				If loRS.Eof Then Exit For
				' 将RS各栏转换为标签值存入标签字典中
				For lvFieldNum = 0 To loRS.Fields.Count - 1
					lvTagName = pBlockName & "." & loRS.Fields(lvFieldNum).Name	' 定义栏名称
					lvTagNum = pBlockName & "." & lvFieldNum + 1	' 定义栏序号
					lvFieldValue = loRS.Fields(lvFieldNum).Value	' 取栏值
					' 第一次先检查有没有单独绑定过标签值,有则后继都使用单独绑定的值,否则使用RS中的栏位值
					If (coTagData.exists(lvTagName) Or coTagData.exists(lvTagNum)) And lvRowNum=1 Then
						lvHasTags(lvFieldNum) = True
					End If
					' 更新之前没有单独绑定过标签值的栏位
					If Not lvHasTags(lvFieldNum) Then
						' 将RS的栏位转换成:"块.栏位名"及"块.栏位序号",分别存入标签字典,以便后继引用
						Call LoadData_(lvTagName, lvFieldValue)
						Call LoadData_(lvTagNum, lvFieldValue)
					End If
				Next
				' 取得完整的块数据内容
				lvBlockData = lvBlockData & GetBlockData_(pBlockName)
				loRS.MoveNext
			Next
		Else
			' 非RS块时,如果已定义块值,则替换掉原块中的内容,可达到动态标签的功能
			If coTagData.Exists(pBlockName) Then
				coBlockBodyMark.Remove(pBlockName)
				coBlockBodyMark.Add pBlockName, coTagData(pBlockName)
			End If
			' 组构块中的所有标签数据
			lvBlockData = lvBlockData & GetBlockData_(pBlockName)
		End If
		' 将更新组构后的块数据存入字典:coBlockData
		If coBlockData.Exists(pBlockName) Then coBlockData.Remove pBlockName
		coBlockData.Add pBlockName, Cstr(lvBlockData)
	End Sub	

	' 组构替换所有标签数据
	' 说明: 1.先更新标签数据, 2.后更新块数据
	Public Sub Build()
		' 1. 更新标签数据
		Call BuildTags_
		' 2. 更新块数据
		Call BuildBlocks_
	End Sub

	' 显示替换完成后的html内容
	Public Sub Show()
		response.write GetHtml
	End Sub


'================================================================================
'== Private
'================================================================================


'--------------------------------------------------------------------------------
'-- Function
'--------------------------------------------------------------------------------

	' 更新模板内容中的所有块及标签内容,并存入字典
	' 调用: SetBlocksAll_(模板文件内容)
	' 返回: 替换成新块标签的块内容
	' 说明: 将块内容存入字典,并将块内容替换为新标签(新标签的作用是为了提高后继搜寻块的速度)
	Private Function SetBlocksAll_(ByVal pHtml)
		Dim lvRule, lvBlockBody, lvBlockName
		Dim loMatches, loMatch
		Dim lvBlockBodyMark, lvBlockBodyTag
		' 取得块内容集的规则(此为闭合块标签规则,块一定要是闭合的才会匹配得到)
		lvRule = "(<!--[\s]*)?(" & "\{" & ")#:(" & ".+?" & ")(" & "}" & ")([\s]*-->)?([\s\S]+?)(<!--[\s]*)?\2/#:\3\4([\s]*-->)?"
		' 取得匹配的块内容集
		Set loMatches = Ebody.RegMatch(pHtml, lvRule)
		' 取得各匹配块的内容
		For Each loMatch In loMatches
			' 取得块名
			lvBlockName = loMatch.SubMatches(2)
			' 取得原始块内容,如:{A.title} a--b {A.addtime}
			lvBlockBody = loMatch.SubMatches(5)	' (5)是指取规则中的第5个索引位所匹配到的内容,从0开始算
			' 取得块内容,包含标签
			lvBlockBodyTag = loMatch.Value
			'----------------------------------------
			' 递归调用,取得子块内容(块里的子块是已经被替换成新标签的内容,这样可以由里到外进行替换为新块标识)
			' 从最里面的块逐层向外替换
			lvBlockBodyMark = SetBlocksAll_(lvBlockBody)
			'----------------------------------------
			' 将原始块内容存入字典:coBlockBody(此块内容包含子块的原始内容)
			If coBlockBody.Exists(lvBlockName) Then coBlockBody.Remove(lvBlockName)
			coBlockBody.Add lvBlockName, Cstr(lvBlockBody)
			' 将替换成新块标签的块内容存入字典:coBlockBodyMark(从里到外存储,如果此块内容包含子块,则子块是经过替换后的块标签)
			If coBlockBodyMark.Exists(lvBlockName) Then coBlockBodyMark.Remove(lvBlockName)
			coBlockBodyMark.Add lvBlockName, lvBlockBodyMark
			' 将当前块内容(包含块标签)替换成新的块标签(从最下层的块向最上层块层层更新)
			' 即去掉原有的块标签及块内容,替换成新的块标签
			' 说明: 这里便于分析和debug数据,暂用---来做为新的标记,如用于应用中,可换为Chr(0)
			pHtml = Replace(pHtml, lvBlockBodyTag, "---" & lvBlockName & "---")
			' 还原块两头的块标签,以便替换掉模板(cvHtml)中的相应块内容
			' 说明: 如果不要这句代码也OK,只不过最终生成的cvHtml的结果中最后一个块内容是展开的.
			'		使用了此名代码,则结果更美观,更易查看结构
			lvBlockBodyMark = Replace(lvBlockBodyTag, lvBlockBody, lvBlockBodyMark)
			' 将cvHtml(模板)内容中的块内容更新为新标签块(从最下层的块向最上层块层层更新)
			cvHtml = Replace(cvHtml, lvBlockBodyMark, "---" & lvBlockName & "---")
			' 返回替换后的内容
			SetBlocksAll_ = pHtml
		Next
		' 如里子块没有匹配的内容,则返回原来的块内容
		If IsEmpty(lvBlockBodyMark) then SetBlocksAll_ = pHtml
	End Function

	' 取得替换后的块内容
	' 调用: GetBlockData_(块标记名)
	' 返回: 更新后的块内容
	' 说明: 此方法用于更新指定块内的所有标签值,并返回带实值的块内容
	Private Function GetBlockData_(ByVal pBlockName)
		Dim lvBlockBody
		' 从新分析后的标签块字典中取得块内容
		lvBlockBody = coBlockBodyMark.item(pBlockName)		
		' 对块内的所有标签进行替换处理
		Call RepTagsData_(lvBlockBody)
		' 返回
		GetBlockData_ = lvBlockBody
	End Function

	' 取得已填充后的块内所有数据(同时返回子块下的所有填充后的内容)
	' 取得指定块下的所有填充后的块内容
	' 调用: Call GetBlockDataAll_(块标记名)
	' 参考: 功能类似于BuildBlocks_
	' 说明: 必须先执行过UpdateBlock才会生效
	Private Function GetBlockDataAll_(ByRef pBlockData)
		Dim lvRule, lvBlockMark, lvBlockName, lvBlockData
		Dim loMatch, loMatches
		' 替换块内容
		' 注意: 不能有两个连续以上的*号,否则报错--未预期的次数符号
		lvRule = "---" & "(\w+?)" & "---"
		' 取得新模板内容(cvHtml)中的所有块标签
		Set loMatches = Ebody.RegMatch(pBlockData, lvRule)
		' 更新模板内所有块标签(从外到里执行替换)
		For Each loMatch In loMatches
			' 取得块名
			lvBlockName = loMatch.SubMatches(0)
			' 取得块标签
			lvBlockMark = loMatch.Value
			' 取块数据(块有更新取更新后的块coBlockData,否则使用原始块内容coBlockBodyMark)
			If coBlockData.Exists(lvBlockName) Then
				lvBlockData = coBlockData.item(lvBlockName)
			Else
				lvBlockData = coBlockBodyMark.item(lvBlockName)
			End If
			' 将cvHTML中的块标签替换为块数据
			pBlockData = Replace(pBlockData, lvBlockMark, lvBlockData)
			' 循环替换填充数据,层层展开,层层替换,直到还原到最里层的块
			pBlockData = GetBlockDataAll_(pBlockData)
		Next
		' 返回
		GetBlockDataAll_ = pBlockData
	End Function


'--------------------------------------------------------------------------------
'-- Sub
'--------------------------------------------------------------------------------

	' 载入模板文件到内存
	' 调用: LoadTpl_(模板文件[可含路径])
	Private Sub LoadTpl_(ByVal pFilePath)
		' 载入FSO类
		Ebody.Use "fso"
		' 无限级载入文件内容到内存中
		cvHtml = Ebody.FSO.GetFileAll(pFilePath)
		' 关闭FSO类
		'Ebody.Close "fso"
		' 将文件中的各块内容存入字典
		Call SetBlocksAll_(cvHtml)
	End Sub

	' 载入文本字符串模板到内存
	' 调用: LoadStr_(字符串)
	Private Sub LoadStr_(ByVal pStr)
		cvHtml = pStr
		' 将文件中的各块内容存入字典
		Call SetBlocksAll_(cvHtml)
	End Sub

	' 载入标签相应的数据源,指定各标签的值,存入字典:coTagData
	' 调用: LoadData_(标签名, 标签值[传引用,对于大数据集速度快])
	' 样例: LoadData_("t_user", oRS),此例会自动将oRS数据源中当前行的所有栏位名及值存入字典中
	' 说明: 此方法用于设定各标签的值,
	'		如果标签名是块标签,则标签值可以是数据对像
	'		如果标签名是纯标签名,则标签值为单个值
	Private Sub LoadData_(ByVal pTagName, ByRef pValue)
		' 记录最新的值,将数据存入对应的标签字典中
		' 注: 如果pValue是数据集,用coTagData其中的一个以块名为标识的item临时存储RS数据,以便后继更新块内容时引用
		If coTagData.Exists(pTagName) Then coTagData.Remove pTagName
		coTagData.Add pTagName, pValue
	End Sub

	' 替换所有独立标签值
	' 调用: Call RepTagsData_(将要替换标签值的内容)
	' 说明: 参数为引用类型,将替换后的字符作为返回
	Private Sub RepTagsData_(ByRef pStr)
		Dim loMatches, loMatch
		Dim lvTagName, lvTagMark, lvBlockBody, lvRule, lvTagData
		' 替换所有独立标签值
		lvRule = "\{(.+?)}"
		Set loMatches = Ebody.RegMatch(pStr, lvRule)
		' 更新所有标签
		For Each loMatch In loMatches
			' 取得标签中的名称(不包含分隔符号),标签名,如: A.title或A.addtime
			lvTagName = loMatch.SubMatches(0)
			' 取得标签(包含分隔符号),用于后继将内容中的标签替换成值,标签内容,如: {A.title}或{A.addtime}
			lvTagMark = loMatch.Value
			' 将标签替换为块数据
			If coTagData.Exists(lvTagName) Then				
				' 取得标签值
				lvTagData = coTagData.Item(lvTagName)				
				' 将所有标签内容替换成原先设定在标签字典中的实值
				If Not IsObject(lvTagData) Then
					pStr = Replace(pStr, lvTagMark, lvTagData)
				End If
			Else
				' 标签未定义时的处理
				Select Case cvUnDefine
				Case "remove"
				pStr = Replace(pStr, lvTagMark, "")
				Case "comment"
				pStr = Replace(pStr, lvTagMark, "<!-- UnDefine Tag "&lvTagName&" -->")
				Case "keep"
				' 保持原样
				End Select
			End If
		Next
	End Sub

	' 填充模板中的所有标签数据
	' 将匹配标签替换为值
	' 调用: Call BuildTags_
	' 说明: 此方法用于将分析后模板中的独立标签进行更新,在模板更新中,优先替换块外的所有独立标签值
	Private Sub BuildTags_()
		' 对内容中的的所有标签进行替换处理
		Call RepTagsData_(cvHtml)
	End Sub

	' 填充模板中的所有块标签数据
	' 将匹配标签替换为值
	' 调用: Call BuildBlocks_
	' 说明: 此方法用于将分析后的模板中的所有块进行分解更新.模板更新时,先更新标签,后替换块内容.
	Private Sub BuildBlocks_()
		Dim lvRule, lvBlockMark, lvBlockName, lvBlockData
		Dim loMatch, loMatches
		' 替换块内容
		' 注意: 不能有两个连续以上的*号,否则报错--未预期的次数符号
		lvRule = "---" & "(\w+?)" & "---"	
		' 取得新模板内容(cvHtml)中的所有块标签
		Set loMatches = Ebody.RegMatch(cvHtml, lvRule)
		' 更新模板内所有块标签(从外到里执行替换)
		For Each loMatch In loMatches
			' 取得块名
			lvBlockName = loMatch.SubMatches(0)
			' 取得块标签
			lvBlockMark = loMatch.Value
			' 取块数据(块有更新取更新后的块coBlockData,否则使用原始块内容coBlockBodyMark)
			If coBlockData.Exists(lvBlockName) Then
				lvBlockData = coBlockData.item(lvBlockName)
			Else
				lvBlockData = coBlockBodyMark.item(lvBlockName)
			End If
			' 将cvHTML中的块标签替换为块数据
			cvHtml = Replace(cvHtml, lvBlockMark, lvBlockData)
			' 循环替换填充数据,层层展开,层层替换,直到还原到最里层的块
			BuildBlocks_()
		Next
	End Sub

	' 绑定tpl内容中的所有块内标签值
	' 调用: Call UpdateBlocks
	' 说明: 此方法用于一次性绑定模板中的所有块内标签值到coTagData字典,以便后继填充内容时使用
	Private Sub UpdateBlocks_()
		' 循环读取字典中的所有块标签
		For Each lvBlockName In coBlockBody
			' 绑定块中的标签值
			UpdateBlock(lvBlockName)
		Next
	End Sub

End Class
%>