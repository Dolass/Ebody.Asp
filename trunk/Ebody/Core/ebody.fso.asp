<%
'################################################################################
'## ebody.fso.asp
'## -----------------------------------------------------------------------------
'## 功能:	FSO文件操作
'## 版本:	1.1 Build 130324
'## 作者:	Tony
'## 日期:	2013/07/02
'## 说明:	Ebody基类
'################################################################################

Class ebody_fso
	
'================================================================================
'== Variable
'================================================================================


'--------------------------------------------------------------------------------
'-- Public
'--------------------------------------------------------------------------------


'--------------------------------------------------------------------------------
'-- Private
'--------------------------------------------------------------------------------
	
	Private cvFsoName		' 文件操作类名
	Private cvCharSet		' 文件字符集
	Private cvForce			' 是否允许删除只读文件
	Private cvOverWrite		' 文件是否允许进行重写操作
	Private cvSizeFormat	' 文件大小显示格式
	Private coFso			' 文件操作类实例

'================================================================================
'== Event
'================================================================================

	Private Sub Class_Initialize()
		'On Error Resume Next
		cvSizeFormat = "K"
		cvOverWrite = True
		cvForce = True
		cvFsoName = "Scripting.FileSystemObject"
		cvCharset = Ebody.CharSet	' "UTF-8"
		Set coFso = Server.CreateObject(cvFsoName)
	End Sub

	Private Sub Class_Terminate()
		'On Error Resume Next
		Set coFso = Nothing
	End Sub

'================================================================================
'== Public
'================================================================================


'--------------------------------------------------------------------------------
'-- Let Property
'--------------------------------------------------------------------------------

	' 设置FSO组件名称
	' 调用: FsoName = 文件操作组件名
	' 样例: Ebody.fso.FsoName = "Scripting.FileSystemObject"
	Public Property Let FsoName(Byval pStr)
		Set coFso = Server.CreateObject(pStr)
	End Property

	' 设置文件编码
	' 调用: CharSet = 文件字符编码
	' 样例: Ebody.fso.CharSet = "UTF-8"
	Public Property Let CharSet(Byval pStr)
		cvCharset = Ucase(pStr)
	End Property

	' 设置是否强制删除只读文件
	' 调用: Force = TRUE/FALSE
	' 样例: Ebody.fso.Force = true
	Public Property Let Force(Byval pBool)
		cvForce = pBool
	End Property

	' 设置是否覆盖原有文件
	' 调用: OverWrite = TRUE/FALSE
	' 样例: Ebody.fso.OverWrite = true
	Public Property Let OverWrite(Byval pBool)
		cbOverwrite = pBool
	End Property

	' 设置文件大小显示格式
	' 调用: SizeFormat = G/M/K/b/auto
	' 样例: Ebody.fso.SizeFormat = K
	Public Property Let SizeFormat(Byval pStr)
		cvSizeFormat = pStr
	End Property

'--------------------------------------------------------------------------------
'-- Get Property
'--------------------------------------------------------------------------------

	' 取得当前页面的物理文件地址
	' 调用: ThisFilePath
	' 返回: H:\WEB\Ebody110\demo.asp
	Public Property Get ThisFilePath()
		ThisFilePath = Request.ServerVariables("Path_Translated")
	End Property

	' 取得当前页面的物理目录地址
	' 调用: ThisDirPath
	' 返回: 页面文件存放于H:\WEB\Ebody110\demo.asp, 返回H:\WEB\Ebody110
	Public Property Get ThisDirPath()
		ThisDirPath = Server.Mappath(".")
		'ThisDirPath = Left(Path, InStrRev(Path, "\"))
	End Property

	' 取得当前页面基于网站主页的相对网络目录路径
	' 调用: RPath
	' 返回: 网站主页 http://localhost, 当前页面地址 http://localhost/tt/tony/demo/home.asp, 返回: /tt/tony/demo
	' 说明: 也就是取当前文件的相对目录地址
	Public Property Get RPath()
		Dim lvPath : lvPath = Request.ServerVariables("Script_Name")
		RPath = Left(lvPath, InStrRev(lvPath, "/") - 1)
	End Property

	' 取得当前页面的文件名(包含扩展名)
	' 调用: ThisFile
	' 返回: 如正浏览的是http://localhost/demo/Index.asp, 则返回Index.asp
	Public Property Get ThisFile()
		Dim lvFileURL : lvFileURL = Request.ServerVariables("PATH_INFO")
		ThisFile = Mid(lvFileURL, InstrRev(lvFileURL, "/", -1, 0) + 1)
	End Property

	' 取得当前文件名(不含扩展名)
	' 调用: ThisFileName
	' 返回: 如文件是Index.asp, 则返回Index
	Public Property Get ThisFileName()
		ThisFileName = Left(ThisFile, Instr(ThisFile, ".") - 1)
	End Property

	' 取得当前文件的扩展名
	' 调用: ThisFileExt
	' 返回: 如文件是Index.asp, 则返回asp
	Public Property Get ThisFileExt()
		ThisFileExt = Mid(ThisFile, InstrRev(ThisFile, ".", -1, 0) + 1)
	End Property

'--------------------------------------------------------------------------------
'-- Function
'--------------------------------------------------------------------------------

	'-------------------------
	' 基础功能类函数
	'-------------------------

	' 创建新对像操作类实例
	' 说明: 当需要创建一个新的实例时可使用此方法
	' 样例: Set loFso = ebody.fso.New
	Public Function [New]()
		Set [New] = New ebody_fso
	End Function

	' 自动命名 (来自Easp2.2 fso)
	' 调用: AutoName()
	' 返回: 随机文件名
	' 说明: 规则是以年月日+时分秒+4位随机号组成
	Public Function AutoName()
		Dim y, m, d, h, mm, S, r
		Randomize
		y = Year(Now)
		m = Month(Now): If m < 10 Then m = "0" & m
		d = Day(Now): If d < 10 Then d = "0" & d
		h = Hour(Now): If h < 10 Then h = "0" & h
		mm = Minute(Now): If mm < 10 Then mm = "0" & mm
		S = Second(Now): If S < 10 Then S = "0" & S
		r = 0
		r = CInt(Rnd() * 1000)
		If r < 10 Then r = "00" & r
		If r < 100 And r >= 10 Then r = "0" & r
		AutoName = y & m & d & h & mm & S & r
	End Function

	' 创建文件夹 (来自Easp2.2 fso)
	' 调用: CreateFolder(文件路径)
	' 样例: CreateFolder("abc/efg/")
	' 返回: True 成功, False 失败
	' 说明: 可支持多层文件夹创建,注意繁简字符
	Public Function CreateFolder(ByVal pFolderPath)
		'On Error Resume Next
		Dim p,arrP,i : CreateFolder = True
		p = Ebody.MapPath(pFolderPath)
		arrP = Split(p,"\") : p = ""
		For i = 0 To Ubound(arrP)
			p = p & arrP(i) & "\"
			If Not IsFolder(p) And i>0 Then coFso.CreateFolder(p)
		Next
		' 返回值,记录错误
		CreateFolder = Not Ebody.IsErr
	End Function

	' 将二进制数据保存为文件
	' 调用: SaveAs(文件路径[支持相对与绝对], 文件数据[二进制])
	' 返回: True 成功, False 失败
	' 说明: 保存数据文件，如：图片，音乐等，可自定义创建在多层文件夹下,没有的文件夹会新增
	Public Function SaveAs(ByVal pFilePath, ByVal pFileData)
		'On Error Resume Next
		' 流模式写入
		pFilePath = Ebody.MapPath(pFilePath)
		SaveAs = CreateFileByStream_(pFilePath, pFileData, 1)	' 写入模式[1 二进制模式/2 文本模式]
	End Function

	' 将文本数据保存为文件
	' 调用: CreateFile(文件路径[支持相对与绝对], 文件数据[文本])
	' 返回: True 成功, False 失败
	' 说明: 保存为文本文件，如：txt，html档等，可自定义创建在多层文件夹下,没有的文件夹会新增
	Public Function CreateFile(ByVal pFilePath, ByVal pFileData)
		'On Error Resume Next
		' 流模式写入
		pFilePath = Ebody.MapPath(pFilePath)
		CreateFile = CreateFileByStream_(pFilePath, pFileData, 2)	' 写入模式[1 二进制模式/2 文本模式]
	End Function

	' 将内容存为文本文件
	' 调用: CreateTextFile(文件路径[支持相对与绝对], 文件数据[txt字串])
	' 返回: True 成功, False 失败
	' 说明: 可自定义创建在多层文件夹下,没有的文件夹会新增 
	Public Function CreateTextFile(ByVal pFilePath, ByVal pFileData)
		'On Error Resume Next
		' 生成空文本文件
		CreateTextFile = CreateFileByStream_(pFilePath, pFileData, 2)
		' 向文本文件中写入内容		
		CreateTextFile = FsoWrite_(pFilePath, pFileData, 1)	' 复写	
	End Function

	' 向文本文件追加内容 (来自Easp2.2 fso)
	' 说明: 只能在文本文件内容后继续追加文本
	Public Function AppendTextFile(ByVal pFilePath, ByVal pFileData)
		pFilePath = Ebody.MapPath(pFilePath)
		AppendTextFile = FsoWrite_(pFilePath, pFileData, 2) ' 追加
	End Function

	' 列出文件夹下的所有文件夹或文件(来自Easp2.2 fso)
	' 调用: List(目录路径, 清单对像类型[file文件/folder目录])
	' 返回: 返回一个二维数组
	Public Function List(ByVal pFolderPath, ByVal pFileType)
		'On Error Resume Next
		Dim f,fs,k,arr(),i,l
		pFolderPath = Ebody.MapPath(pFolderPath) : i = 0
		Select Case LCase(pFileType)
			Case "0","" l = 0
			Case "1","file" l = 1
			Case "2","folder" l = 2
			Case Else l = 0
		End Select
		Set f = coFso.GetFolder(pFolderPath)
		If l = 0 Or l = 2 Then
			Set fs = f.SubFolders
			ReDim Preserve arr(4,fs.Count-1)
			For Each k In fs
				arr(0,i) = k.Name & "/"
				arr(1,i) = FormatSize_(k.Size,cvSizeFormat)
				arr(2,i) = k.DateLastModified
				arr(3,i) = Attr2Str_(k.Attributes)
				arr(4,i) = k.Type
				i = i + 1
			Next
		End If
		If l = 0 Or l = 1 Then
			Set fs = f.Files
			ReDim Preserve arr(4,fs.Count+i-1)
			For Each k In fs
				arr(0,i) = k.Name
				arr(1,i) = FormatSize_(k.Size,cvSizeFormat)
				arr(2,i) = k.DateLastModified
				arr(3,i) = Attr2Str_(k.Attributes)
				arr(4,i) = k.Type
				i = i + 1
			Next
		End If
		Set fs = Nothing
		Set f = Nothing
		List = arr
		' 记录错误
		Ebody.Log("List")
	End Function

	' 重命名文件或文件夹(来自Easp2.2 fso)
	' 返回: true 成功, false 失败
	Public Function Rename(ByVal pPath, ByVal pNewName)
		Dim lvNew
		pPath = Ebody.MapPath(pPath)
		Rename = True
		lvNew = Left(pPath,InstrRev(pPath,"\")) & pNewName
		' 文件或文件夹是否存在
		If Not IsExists(pPath) Then
			Rename = False
'			' 记录错误
			Ebody.Log("Rename")
			Exit Function
		End If
		' 如果新文件存在,则不更新
		If IsExists(lvNew) Then
			Rename = False
'			' 记录错误
			Ebody.Log("Rename")
			Exit Function
		End If
		' 用移动操作来模拟重命名操作
		If IsFolder(pPath) Then			
			coFso.MoveFolder pPath,lvNew
		ElseIf IsFile(pPath) Then
			coFso.MoveFile pPath,lvNew
			'Copy pPath, lvNew
			'Del pPath
		End If
	End Function

	' 删除文件（支持通配符*和?）(来自Easp2.2 fso)
	' 返回: true 成功, false 失败
	Public Function DelFile(ByVal pPath)
		DelFile = FOFO(pPath,"",0,2)
	End Function

	' 删除文件夹（支持通配符*和?）(来自Easp2.2 fso)
	' 返回: true 成功, false 失败
	Public Function DelFolder(ByVal pPath)
		DelFolder = FOFO(pPath,"",1,2)
	End Function

	' 删除文件或文件夹(来自Easp2.2 fso)
	' 返回: true 成功, false 失败
	' 说明: 此方法可删除服务器上的单个文件夹及文件, 不支持通配符*和?
	Public Function Del(ByVal pPath)
		pPath = Ebody.MapPath(pPath)
		If IsFile(pPath) Then
			Del = DelFile(pPath)
		ElseIf IsFolder(pPath) Then
			Del = DelFolder(pPath)
		Else
			Del = False'			
		End If
		' 记录错误
		Ebody.Log("Del")
	End Function

	' 移动文件 (支持通配符*和?) (来自Easp2.2 fso)
	' 返回: true 成功, false 失败
	' 说明: 支持通配符*和?
	Public Function MoveFile(ByVal pFromPath, ByVal pToPath)
		MoveFile = FOFO(pFromPath,pToPath,0,1)
	End Function

	' 移动文件夹 (支持通配符*和?) (来自Easp2.2 fso)
	' 返回: true 成功, false 失败
	' 说明: 支持通配符*和?
	Public Function MoveFolder(ByVal pFromPath, ByVal pToPath)
		MoveFolder = FOFO(pFromPath,pToPath,1,1)
	End Function

	' 移动文件或文件夹 (来自Easp2.2 fso)
	' 返回: true 成功, false 失败
	' 说明: 此方法可移动服务器上的单个文件夹及文件, 不支持通配符*和?
	Public Function Move(ByVal pFromPath, ByVal pToPath)
		Dim ff,tf : ff = Ebody.MapPath(pFromPath) : tf = Ebody.MapPath(pToPath)
		If IsFile(ff) Then
			Move = MoveFile(pFromPath,pToPath)
		ElseIf IsFolder(ff) Then
			Move = MoveFolder(pFromPath,pToPath)
		Else
			Move = False
		End If
		' 记录错误
		Ebody.Log("Move")
	End Function

	' 复制文件 (支持通配符*和?) (来自Easp2.2 fso)
	' 返回: true 成功, false 失败
	Public Function CopyFile(ByVal pFromPath, ByVal pToPath)
		CopyFile = FOFO(pFromPath,pToPath,0,0)
	End Function

	' 复制文件夹 (支持通配符*和?)(来自Easp2.2 fso)
	' 返回: true 成功, false 失败
	Public Function CopyFolder(ByVal pFromPath, ByVal pToPath)
		CopyFolder = FOFO(pFromPath,pToPath,1,0)
	End Function

	' 复制文件或文件夹(来自Easp2.2 fso)
	' 返回: true 成功, false 失败
	' 说明: 此方法可复制服务器上的单个文件夹及文件, 不支持通配符*和?
	Public Function Copy(ByVal pFromPath, ByVal pToPath)
		Dim ff,tf : ff = Ebody.MapPath(pFromPath) : tf = Ebody.MapPath(pToPath)
		If IsFile(ff) Then
			Copy = CopyFile(pFromPath,pToPath)
		ElseIf IsFolder(ff) Then
			Copy = CopyFolder(pFromPath,pToPath)
		Else
			Copy = False
		End If
		' 记录错误
		Ebody.Log("Copy")
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
		'If IsFile(pPath) Or IsFolder(pPath) Or IsDrive(pPath) Then IsExists = True Else IsExists = False
		IsExists = Ebody.IsExists(pPath)
	End Function

	' 检测文件是否存在
	' 调用: IsFile(文件路径[支持相对与绝对])
	' 返回: true 存在, false 不存在
	' 样例: IsFile("/dev/demo/abc.asp")
	' 说明: 地址第一个符号是/, 则位置从根目录开始(即:绝对地址), 不带则位置从当前文件所在文件夹开始(即:相对地址)
	Public Function IsFile(ByVal pFilePath)
		'pFilePath = Ebody.MapPath(pFilePath)
		'If coFso.FileExists(pFilePath) Then IsFile = True Else IsFile = False
		IsFile = Ebody.IsFile(pFilePath)
	End Function
	
	' 检测文件夹是否存在
	' 调用: IsFolder(目录路径[支持相对与绝对])
	' 返回: true 存在, false 不存在
	' 样例: IsFolder("/common/system")
	' 说明: 地址第一个符号是/, 则位置从根目录开始(即:绝对地址), 不带则位置从当前文件所在文件夹开始(即:相对地址)
	Public Function IsFolder(ByVal pFolderPath)
		'pFolderPath = Ebody.MapPath(pFolderPath)
		'If coFso.FolderExists(pFolderPath) Then IsFolder = True Else IsFolder = False
		IsFolder = Ebody.IsFolder(pFolderPath)
	End Function

	' 检测驱动器是否存在
	' 调用: IsDrive(盘符)
	' 返回: true 存在, false 不存在
	' 样例: IsDrive("d:")
	Public Function IsDrive(ByVal pDrive)
		'pDrive = Ebody.MapPath(pDrive)
		'If coFSO.DriveExists(pDrive) Then IsDrive = True Else IsDrive = False
		IsDrive = Ebody.IsDrive(pDrive)
	End Function

	'-------------------------
	' 系统取值类函数(Get打头)
	'-------------------------

	' 读取指定文本文件内容
	' 调用: GetFile(文件路径[支持相对与绝对])
	' 返回: 文件字符串内容
	' 样例: GetFile("abc/Readme.txt")
	' 说明: 读取当前文件内容, 通过数据流控件(ADODB.Stream)读取或FSO读取, 只能读取文本文件
	Public Function GetFile(ByVal pFilePath)
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
					lvFileData = Ebody.RegReplace(lvFileData, "^\uFEFF", "")
				End If
			Case "add"
				If Not Test(lvContent, "^\uFEFF") Then
					lvFileData = Chrw(&hFEFF) & lvFileData
				End If
			End Select
		End If
		' 返回
		GetFile = lvFileData
	End Function

	' 无限级读取文件内容, 同时也读取所有包含文件的内容
	' 调用: GetFileAll(文件路径[支持相对与绝对])
	' 返回: 所有Include中的文件的内容
	' 说明: 此方法支持带有<inclued>标签的文件内容读取
	Public Function GetFileAll(ByVal pFilePath)
		Dim lvContent, lvRule, lvIncFile, lvIncStr
		Dim loMatch, loMatches
		' 读取当前文件的内容
		pFilePath = Ebody.MapPath(pFilePath)
		lvContent = GetFile(pFilePath)
		If IsNull(lvContent) Then Exit Function
		' 替换为特定标签
		lvContent = Ebody.RegReplace(lvContent, "<% *?@.*?%"&">", "")
		lvContent = Ebody.RegReplace(lvContent, "(<%[^>]+?)(option +?explicit)([^>]*?%"&">)", "$1'$2$3")
		' 定义搜寻包含文件的规则
		lvRule = "<!-- *?#include +?(file|virtual) *?= *?""??([^"":?*\f\n\r\t\v]+?)""?? *?-->"		
		' 验证内容是否有效, 并进行读取
		If Ebody.RegTest(lvContent, lvRule) Then
			Set loMatches = Ebody.RegMatch(lvContent, lvRule)
			For Each loMatch In loMatches
				If LCase(loMatch.SubMatches(0)) = "virtual" Then
					lvIncFile = loMatch.SubMatches(1)
				Else
					lvIncFile = Mid(pFilePath, 1, InstrRev(pFilePath,Ebody.IIF(Instr(pFilePath, ":")>0, "\", "/"))) & loMatch.SubMatches(1)
				End If
				lvIncStr = GetFileAll(lvIncFile)
				lvContent = Replace(lvContent, loMatch, lvIncStr)
			Next
			Set loMatches = Nothing
		End If
		GetFileAll = lvContent
	End Function

	' 获取文件对像属性信息
	' 调用: GetAttr(对像路径, 属性类别)
	' 样例: GetAttr("common/system","size")
	' 说明: 分为三种类型,文件,文件夹,驱动器.
	'		属性分别以不同类型来使用
	Public Function GetAttr(ByVal pPath, ByVal pAttrType)
		'On Error Resume Next
		Dim loResource, lvAttrValue 
		pPath = Ebody.MapPath(pPath)
		' 判断类型并取得资源引用
		If IsDrive(pPath) Then
			Set loResource = coFso.GetDrive(pPath)	' 驱动器
		ElseIf IsFolder(pPath) Then
			Set loResource = coFso.GetFolder(pPath)	' 文件夹
		ElseIf IsFile(pPath) Then
			Set loResource = coFso.GetFile(pPath)	' 文件			
		Else
			GetAttr = ""
			Exit Function
		End If
		' 依属性类别取得资源属性
		Select Case LCase(pAttrType)
			' 文件/文件夹属性
			Case "0","name" : lvAttrValue = loResource.Name	' 资源名称
			Case "1","date", "datemodified" : lvAttrValue = loResource.DateLastModified	' 上次修改时间
			Case "2","datecreated" : lvAttrValue = loResource.DateCreated	' 创建时间
			Case "3","dateaccessed" : lvAttrValue = loResource.DateLastAccessed	' 上次访问时间
			Case "4","size" : lvAttrValue = FormatSize_(loResource.Size,cvSizeFormat)	' 资源大小
			Case "5","attr" : lvAttrValue = Attr2Str_(loResource.Attributes)	' 资源属性(如:只读,可写等)
			Case "6","type" : lvAttrValue = loResource.Type	' 资源类型
			Case "7","path" : lvAttrValue = loResource.Path	' 资源物理路径
			Case "8","parentfolder" :  lvAttrValue = loResource.ParentFolder	' 父文件夹
			' 驱动器属性
			Case "9","rootfolder" :  lvAttrValue = loResource.RootFolder	' 根文件夹
			Case "10","volumename" : lvAttrValue = loResource.VolumeName	' 设置或返回指定驱动器的卷标名
			Case "11","totalsize" : lvAttrValue = FormatSize_(loResource.TotalSize, cvSizeFormat)	' 驱动器或网络共享的总空间大小
			Case "12","freespace" : lvAttrValue = FormatSize_(loResource.FreeSpace, cvSizeFormat)	' 指定驱动器上或共享驱动器可用的磁盘空间
			Case "13","availablespace" : lvAttrValue = FormatSize_(loResource.AvailableSpace, cvSizeFormat)	' 指定的驱动器或网络共享上的用户可用的空间容量
			Case "14:","filesystem" : lvAttrValue = loResource.FileSystem	' 取得磁盘格式类型(可用的返回类型包括 FAT、NTFS 和 CDFS)
			Case Else lvAttrValue = ""
		End Select
		' 释放资源对像
		Set loResource = Nothing
		' 返回值
		GetAttr = lvAttrValue
	End Function

	' 取文件名(来自Easp2.2 fso)
	' 样例: GetFileName("common/abc/abc.asp"), 返回abc
	' 返回: 文件名字串
	' 说明: 取一个字串中,以.区格的文件名
	Public Function GetFileName(ByVal pFilePath)
		GetFileName = GetNameOf_(pFilePath, 0)
	End Function

	' 取文件扩展名(来自Easp2.2 fso)
	' 样例: GetFileExt("common/abc/abc.asp"), 返回asp
	' 返回: 扩展名字串
	' 说明: 取一个字串中,以.区格的扩展名
	Public Function GetFileExt(ByVal pFilePath)
		GetFileExt = GetNameOf_(pFilePath, 1)
	End Function

	'-------------------------
	' 系统设定类函数(Set打头)
	'-------------------------

	' 设置文件对像属性(来自Easp2.2 fso)
	Public Function SetAttr(ByVal pPath, ByVal pAttrType)
		'On Error Resume Next
		Dim lvAttrs,lvAttrNum,i,n
		Dim loResource
		pPath = Ebody.MapPath(pPath)
		n = 0
		SetAttr = True
		If Not IsExists(pPath) Then
			SetAttr = False
			Ebody.Log("SetAttr")
			Exit Function
		End If
		' 取得对像资源
		If IsDrive(pPath) Then
			Set loResource = coFso.GetDrive(pPath)	' 驱动器
		ElseIf IsFolder(pPath) Then
			Set loResource = coFso.GetFolder(pPath)	' 文件夹
		ElseIf IsFile(pPath) Then
			Set loResource = coFso.GetFile(pPath)	' 文件		
		Else
			SetAttr = False
			Exit Function
		End If
		' 设定对像属性
		lvAttrNum = loResource.Attributes
		pAttrType = UCase(pAttrType)
		If Instr(pAttrType, "+")>0 Or Instr(pAttrType, "-")>0 Then
			lvAttrs = Ebody.IIF(Instr(pAttrType, " ")>0, Split(pAttrType, " "), Split(pAttrType, ","))
			For i = 0 To Ubound(lvAttrs)
				Select Case lvAttrs(i)
					Case "+R" If lvAttrNum And 1 Then lvAttrNum = lvAttrNum Else lvAttrNum = lvAttrNum + 1
					Case "-R" If lvAttrNum And 1 Then lvAttrNum = lvAttrNum - 1 Else lvAttrNum = lvAttrNum
					Case "+H" If lvAttrNum And 2 Then lvAttrNum = lvAttrNum Else lvAttrNum = lvAttrNum + 2
					Case "-H" If lvAttrNum And 2 Then lvAttrNum = lvAttrNum - 2 Else lvAttrNum = lvAttrNum
					Case "+S" If lvAttrNum And 4 Then lvAttrNum = lvAttrNum Else lvAttrNum = lvAttrNum + 4
					Case "-S" If lvAttrNum And 4 Then lvAttrNum = lvAttrNum - 4 Else lvAttrNum = lvAttrNum
					Case "+A" If lvAttrNum And 32 Then lvAttrNum = lvAttrNum Else lvAttrNum = lvAttrNum + 32
					Case "-A" If lvAttrNum And 32 Then lvAttrNum = lvAttrNum - 32 Else lvAttrNum = lvAttrNum
				End Select
			Next
			loResource.Attributes = lvAttrNum
		Else
			For i = 1 To Len(lvAttrs)
				Select Case Mid(lvAttrs, i, 1)
					Case "R" n = n + 1
					Case "H" n = n + 2
					Case "S" n = n + 4
				End Select
			Next
			loResource.Attributes = Ebody.IIF(lvAttrNum And 32, n + 32, n)
			If lvAttrNum And 32 Then loResource.Attributes = n + 32 Else loResource.Attributes = n
		End If
		Set loResource = Nothing
		' 捕捉错误
		If Err.Number<>0 Then
			SetAttr = False
			Ebody.Log("SetAttr")
		End If
		Err.Clear()
	End Function

'--------------------------------------------------------------------------------
'-- Sub
'--------------------------------------------------------------------------------

	'-------------------------
	' 基础功能类过程
	'-------------------------

	' 自动下载文件
	' 调用: DownloadFile(文件地址[相对或绝对])
	' 说明: 如找不到下载文件,则提示信息
	Public Sub DownloadFile(ByVal pFilePath)
		'On Error Resume Next
		pFilePath = Ebody.MapPath(pFilePath)
		' 属性设定
		Response.Buffer = True
		Response.Clear
		' 流对像
		Dim loStream : Set loStream = Server.CreateObject("ADODB.Stream")  
		loStream.Open   
		loStream.Type = 1   ' 二制制读取
		'On Error Resume Next
		' fso对像
		Dim loFso : Set loFso = Server.CreateObject("Scripting.FileSystemObject")
		' 验证目标文件是否存在,并执行下载
		If Not loFso.FileExists(pFilePath) Then			
			response.write "<script>alert('无法下载，在服务器找不到该文件!')</script>" 
		Else
			' 取得文件资源
			Dim loFile : Set loFile = loFso.GetFile(pFilePath)
			Dim lvFilelength : lvFilelength = loFile.size   
			loStream.LoadFromFile(pFilePath)		
			' 生成并下载
			If err Then
				' 失败,提示信息
				Response.write "<script>alert(" & err.Description & ")</script>"			
			Else
				' 成功,配置下载参数,执行下载
				Response.AddHeader "Content-Disposition","attachment; filename=" & loFile.name  
				Response.AddHeader "Content-Length",lvFilelength
				Response.CharSet = "UTF-8"
				Response.ContentType = "application/octet-stream"
				Response.BinaryWrite loStream.Read	' 生成二进制文件
				Response.Flush
			End If
		End If

		' 关闭对像
		Set loFile = Nothing
		Set loFso = Nothing
		loStream.Close
		Set loStream = Nothing  
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

	' 格式化文件大小 (来自Easp2.2 fso)
	' 调用: FormatSize_(文件大小, 格式化级别[G/M/K/B])
	Private Function FormatSize_(Byval fileSize, ByVal level)
		Dim s : s = Int(fileSize) : level = UCase(level)
		If s/(1073741824)>0.01 Then FormatSize_ = FormatNumber(s/(1073741824),2,-1,0,-1) & " GB" Else FormatSize_ = "0.01" & " GB"

		If s = 0 Then FormatSize_ = "0 GB"
		If level = "G" Or (level="AUTO" And s>1073741824) Then Exit Function
		If s/(1048576)>0.1 Then FormatSize_ = FormatNumber(s/(1048576),1,-1,0,-1) & " MB" Else FormatSize_ = "0.1" & " MB"
		If s = 0 Then FormatSize_ = "0 MB"
		If level = "M" Or (level="AUTO" And s>1048576) Then Exit Function
		If (s/1024)>1 Then FormatSize_ = Int(s/1024) & " KB" Else FormatSize_ = 1 & " KB"

		If s = 0 Then FormatSize_ = "0 KB"
		If Level = "K" Or (level="AUTO" And s>1024) Then Exit Function
		If level = "B" or level = "AUTO" Then
			FormatSize_ = s & " bytes"
		Else
			FormatSize_ = s
		End If
	End Function

	' 取得格式化的文件属性 (来自Easp2.2 fso)
	' 返回格式: 属性代码-属性描述
	Private Function Attr2Str_(ByVal attr)
		Dim a,s : a = Int(attr)
		Dim lvAttrDesc
		If a>=2048 Then a = a - 2048
		If a>=1024 Then a = a - 1024
		If a>=32 Then : s = "A" : a = a- 32 : End If
		If a>=16 Then a = a- 16
		If a>=8 Then a = a - 8
		If a>=4 Then : s = "S" & s : a = a- 4 : End If
		If a>=2 Then : s = "H" & s : a = a- 2 : End If
		If a>=1 Then : s = "R" & s : a = a- 1 : End If

		select Case attr
		Case 0 lvAttrDesc = "普通文件。没有设置任何属性。 "
		Case 1 lvAttrDesc = "只读文件。可读写。 "
		Case 2 lvAttrDesc = "隐藏文件。可读写。 "
		Case 4 lvAttrDesc = "系统文件。可读写。 "
		Case 16 lvAttrDesc = "文件夹或文件夹。只读。 "
		Case 32 lvAttrDesc = "上次备份后已更改的文件。可读写。 "
		Case 1024 lvAttrDesc = "链接或快捷方式。只读。 "
		Case 2048 lvAttrDesc = "压缩文件。只读。"
		Case Else lvAttrDesc = "内容不详"
		End select

		Attr2Str_ = s & "-" & lvAttrDesc
	End Function

	'-------------------------
	' 验证类函数(Is打头)
	'-------------------------

	' 路径是否包含通配符 (参考Easp2.2 fso)
	' 调用: IsWildcards_(路径)
	' 返回: true 包含, false 不包含
	Private Function IsWildcards_(ByVal pPath)
		IsWildcards_ = False
		If Instr(pPath,"*")>0 Or Instr(pPath,"?")>0 Then IsWildcards_ = True
	End Function

	'-------------------------
	' 系统取值类函数(Get打头)
	'-------------------------

	' 读取文件内容方法1 (使用数据流控件ADODB.Stream读取)
	' 调用: GetFileByStream_(文件路径[支持相对与绝对], 读取类型[1 二进制模式/2 文本模式])
	' 返回: 依据读取类型返回, pReadType=1 返回二进制内容, pReadType=2 返回文本内容
	' 说明: 只读取当前文件内容, 不支持多级include读取
	Private Function GetFileByStream_(ByVal pFilePath, ByVal pReadType)
		Dim lvFileData, loStream
		pFilePath = Ebody.MapPath(pFilePath)
		Set loStream = Server.CreateObject("ADODB.Stream")
		' 开始读取
		With loStream
			.Type = pReadType	' 读取模式
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
		pFilePath = Ebody.MapPath(pFilePath)
		' 打开并读取文件内容, 直到文件结尾, 则返回
		Set loFile = coFSO.OpenTextFile(pFilePath, 1, False, -1)
		If loFile.AtEndOfStream = false Then GetFileByFSO_ = loFile.ReadAll
		' 关闭文件
		loFile.Close
		Set loFile=Nothing
	End Function

	' 取文件名或扩展名(来自Easp2.2 fso)
	Private Function GetNameOf_(ByVal f, ByVal t)
		Dim re,na,ex
		If Ebody.IsNull(f) Then GetNameOf_ = "" : Exit Function
		f = Replace(f,"\","/")
		If Right(f,1) = "/" Then
			re = Split(f,"/")
			GetNameOf_ = Ebody.IIF(t=0, re(Ubound(re)-1), "")
			Exit Function
		ElseIf Instr(f,"/")>0 Then
			re = Split(f,"/")(Ubound(Split(f,"/")))
		Else
			re = f
		End If
		If Instr(re,".")>0 Then
			na = Left(re,InstrRev(re,".")-1)
			ex = Mid(re,InstrRev(re,".")+1)
		Else
			na = re
			ex = ""
		End If
		If t = 0 Then
			GetNameOf_ = na
		ElseIf t = 1 Then
			GetNameOf_ = ex
		End If
	End Function

	' 将数据保存为文件 (使用数据流控件ADODB.Stream写入)
	' 调用: CreateFileByStream_(文件路径[支持相对与绝对], 文件数据, 写入模式[1 二进制模式/2 文本模式])
	' 返回: True 成功, False 失败
	' 说明: 可自定义创建在多层文件夹下,没有的文件夹会新增
	Private Function CreateFileByStream_(ByVal pFilePath, ByVal pFileData, ByVal pWriteType)
		'On Error Resume Next
		Dim lvFileData, loStream
		' 是否复写文件(1 不复写, 2 复写)
		Dim lvOverWriteType : If cvOverWrite Then lvOverWriteType = 2 Else lvOverWriteType = 1
		' 得到文件写入路径
		pFilePath = Ebody.MapPath(pFilePath)
		' 建文件夹
		CreateFileByStream_ = CreateFolder(Left(pFilePath,InstrRev(pFilePath,"\")-1))
		' 开始写入
		If CreateFileByStream_ Then
			Set loStream = Server.CreateObject("ADODB.Stream")
			With loStream
				.Type = pWriteType						' 1 以二进制模式写入, 2 文本模式写入
				.Mode = 3								' 可读可写
				.Charset = cvCharset					' 设定字符集
				.Open									' 打开文件对像
				.Position = loStream.Size				' 定位到文件未尾,在文件未尾继续写入				
				If pWriteType = 1 Then
					.Write = pFileData					' 以二进制模式写入
				Else
					.WriteText = pFileData				' 以文本模式写入
				End If				
				.SaveToFile pFilePath,lvOverWriteType	' 生成文件
				.Flush									' 将缓冲区中的Stream数据立即写入
				.Close									' 关闭文件对像
			End With
			Set loStream = Nothing
		End If
		' 捕捉错误
		If Err.Number<>0 Then
			CreateFileByStream_ = False
			Ebody.Log("CreateFileByStream_")
		Else
			CreateFileByStream_ = True
		End If
	End Function

	' 将数据写入文本文件 (使用文件系统控件Fso写入)
	' 调用: FsoWrite_(文件路径[支持相对与绝对], 文件数据, 写入模式[1 复写, 2 追加] )
	' 返回: 文件字符串内容
	' 说明: 只能读取或写入文本文件(txt), 不支持大字符集文件
	Private Function FsoWrite_(ByVal pFilePath, ByVal pFileData, ByVal pWriteType)
		Dim loFso
		pFilePath = Ebody.MapPath(pFilePath)
		' 写入模式
		Select Case pWriteType
		Case 1
			lvOverWriteType = 2	' 文件进行重写操作
		Case 2
			lvOverWriteType = 8	' 在文件末尾进行追加操作
		End Select
		' 写入并生成文件
		Set loFso = coFSO.OpenTextFile(pFilePath, lvOverWriteType, True)		
		loFso.Write pFileData		' 单行写入
		'loFso.WriteLine pFileData	' 换行写入
		' 关闭文件对像
		loFso.Close
		Set loFso = Nothing
		' 捕捉错误
		If Err.Number<>0 Then
			FsoWrite_ = False
			Ebody.Log("FsoWrite_")
		Else
			FsoWrite_ = True
		End If
	End Function

	' 文件或文件夹操作原型 (来自Easp2.2 fso) ####待完善####
	' 调用: FOFO(来源, 目标, 对像类型[文件/目录], 动作[复制/移动/删除])
	Private Function FOFO(ByVal pFromPath, ByVal pToPath, ByVal FOF, ByVal MOC)
		' 错误忽略(在此此功能必须开启)
		On Error Resume Next
		FOFO = True
		Dim ff,tf,oc,of,oi,ot,os
		' 动态组构的是否复写选项
		Dim lvOverWriteStr : If MOC=0 Then lvOverWriteStr = ",cvOverWrite"
		'ff 来源路径				 'tf 目标路径
		ff = Ebody.MapPath(pFromPath) : tf = Ebody.MapPath(pToPath)
		If FOF = 0 Then
		'如果是文件
			oc = IsFile(ff) : of = "File" : oi = "文件"
		ElseIf FOF = 1 Then
		'如果是文件夹
			oc = IsFolder(ff) : of = "Folder" : oi = "文件夹"
		End If
		If MOC = 0 Then
			ot = "Copy" : os = "复制"
		ElseIf MOC = 1 Then
			ot = "Move" : os = "移动"
		ElseIf MOC = 2 Then
			ot = "Delete" : os = "删除"
		End If
		If oc Then
		'如果文件或文件夹存在
			If MOC<>2 Then
			'如果复制和移动
				If FOF = 0 Then
				'如果是文件
					If Right(pToPath,1)="/" or Right(pToPath,1)="\" Then
					'如果目标路径是文件夹，直接建立
						FOFO = CreateFolder(tf) : tf = tf & "\"
					Else
					'如果目标路径是文件，建立文件夹
						FOFO = CreateFolder(Left(tf,InstrRev(tf,"\")-1))
					End If
				ElseIf FOF = 1 Then
				'如果是文件夹则先建立目标文件夹
					tf = tf & "\"
					FOFO = CreateFolder(tf)
				End If
				'执行复制或者移动，如果是复制要考虑是否覆盖
				Execute("coFso."&ot&of&" ff,tf"&lvOverWriteStr)
				'Easp.wn("Fso."&ot&of&" "&ff&","&tf&","&b_overwrite&"")
			Else
				'删除，考虑是否删除只读
				Execute("coFso."&ot&of&" ff,cvforce")
			End If
			If Err.Number<>0 Then
				FOFO = False
				'Easp.Error.Msg = "<br />" & os & oi & "失败！" & "( "&frompath&" "&Easp.IIF(MOC=2,"",os&"到 "&pToPath)&" )"
				'Easp.Error.Raise 63
			End If
		ElseIf isWildcards_(ff) Then
			'如果有通配符
'			If Not IsFolder(Left(ff,InstrRev(ff,"\")-1)) Then
'				FOFO = False
'				Easp.Error.Msg = "<br />" & os & oi & "失败！" & Easp.IIF(MOC=2,"","源") & oi & "不存在( "&frompath&" )"
'				Easp.Error.Raise 63
'			End If
			If MOC<>2 Then
				'复制和移动
				FOFO = CreateFolder(tf)
				Execute("coFso."&ot&of&" ff,tf"&lvOverWriteStr)
			Else
				'删除
				Execute("coFso."&ot&of&" ff,cvforce")
			End If
			If Err.Number<>0 Then
				FOFO = False
				'Easp.Error.Msg = "<br />" & os & oi & "失败！" & "( "&frompath&" "&Easp.IIF(MOC=2,"",os&"到 "&pToPath)&" )"
				'Easp.Error.Raise 63
			End If
		Else
			FOFO = False
			'Easp.Error.Msg = "<br />" & os & oi & "失败！" & Easp.IIF(MOC=2,"","源")&oi&"不存在( "&frompath&" )"
			'Easp.Error.Raise 63
		End If
		Err.Clear()
	End Function

'--------------------------------------------------------------------------------
'-- Sub
'--------------------------------------------------------------------------------

	'-------------------------
	' 基础功能类过程
	'-------------------------

	'-------------------------
	' 设置对像类过程(Set/Remove打头)
	'-------------------------

End Class
%>