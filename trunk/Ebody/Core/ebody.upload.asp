<%
'################################################################################
'## ebody.upload.asp
'## -----------------------------------------------------------------------------
'## 功能:	文件上传类
'## 版本:	1.1 Build 130324
'## 作者:	Tony
'## 日期:	2013/03/24
'## 说明:	基本类
'################################################################################
' 示例分析:
'################################################################################
' 以下列出上传上来的数据内容,用response.BinaryWrite方法打印出来的,供分析
' ←: 代表一个回车符
' ↓: 代表一个换行符
' □: 代表一个空格
' 数据的基本格式是有规则的,所以可以通过相同的规则来分析出各个内容
' 1. 每个数据内容均由一个固定的分隔符字隔开,如:-----------------------------9930139772306
' 2. 文件头与分隔符之前由←↓隔开
' 3. 文件实体数据与文件头由←↓←↓隔开
' 4. 在最后仍会有一个分隔符,注意:此分隔符与先前的有所不同,如:-----------------------------9930139772306--
' 5. 数据分隔符在同一批上传过程中都是一样的,所以我们可以依此来分隔出每个上传对像
' 以下数据示例:
'-----------------------------9930139772306←↓Content-Disposition:□form-data;□name="text"←↓←↓测试文本←↓ '-----------------------------9930139772306←↓Content-Disposition:□form-data;□name="file";□filename="404.gif"←↓'Content-Type:□image/gif←↓←↓.............这里是二进制的数据内容,显示出来的是乱码,而且很多,为方便举列,故省略.............←↓-----------------------------9930139772306--←↓
'################################################################################


Class ebody_upload
	
'================================================================================
'== Variable
'================================================================================


'--------------------------------------------------------------------------------
'-- Public
'--------------------------------------------------------------------------------
	Public File		' 公共的文件
	Public Form		' 公共的表单
	Public SuccessCount	' 上传成功的总数
	Public FailCount	' 上传失败的总数

'--------------------------------------------------------------------------------
'-- Private
'--------------------------------------------------------------------------------
	Private coUpStream		' 上传流对像
	Private coFileStream	' 文件流对像
	Private coFso			' 用于验证服务器fso是否激活
	Private cvCharset		' 字符集
	Private cvAllowed		' 白名单
	Private cvDenied		' 单名单
	Private cvSavePath		' 保存路径
	Private cvFileMaxSize	' 最大文件上传大小
	Private cvTotalMaxSize	' 总上传大小
	Private cvBlockSize		' 块大小
	Private cvAutoMD		' 是否自动创建目录
	Private cvAutoName		' 是否自动文件名
	Private cvForce			' 是否强制替换原有文件
	Private cvErrMsg		' 错误信息


'================================================================================
'== Event
'================================================================================

	Private Sub Class_Initialize()
		'On Error Resume Next

		' 创建及验证组件
		' --------------------------------------------------------------------------
		' 接收上传的数据储存对像
		Set coUpStream = Server.CreateObject("ADODB.Stream")
		' 保存分析后的各数据对像
		Set coFileStream = Server.CreateObject("ADODB.Stream")
		' 验证组件是否安装
		If Err.number<>0 Then cvErrMsg = "创建流对象(ADODB.STREAM)时出错,可能系统不支持或没有开通该组件" : Exit Sub
		' Fso
		Set coFso = Server.CreateObject("Scripting.FileSystemObject")
		' 验证组件是否安装
		If Err.number <> 0 Then cvErrMsg = "创建文件管理对像(FSO)时出错,可能系统不支持或没有开通该组件" : Exit Sub
		' 保存文件信息的字典
		Set File = Server.CreateObject("Scripting.Dictionary")
		' 保存表单值的字典
		Set Form = Server.CreateObject("Scripting.Dictionary")

		' 初始化
		' --------------------------------------------------------------------------
		' 文件上传成功与失败数
		SuccessCount = 0
		FailCount = 0
		' 默认编码
		cvCharset = Ebody.Charset
		' 默认仅允许上传类型,所有文件
		cvAllowed = ""
		' 默认不允许上传类型,没有
		cvDenied = ""
		' 默认文件保存位置,当前目录
		cvSavePath = ""
		' 默认自动建立不存在的文件夹
		cvAutoMD = True
		' 默认不使用随机文件名
		cvAutoName = False
		' 默认不复写存在文件及文件夹
		cvForce = False
		' 默认单个文件允许大小,不限制
		cvFileMaxSize = 0
		' 默认总文件大小,不限制
		cvTotalMaxSize = 0
		' 分块默认每次上传10K
		cvBlockSize = 10 * 1024
	End Sub

	Private Sub Class_Terminate()
		'On Error Resume Next
		' 清除变量及对像
		File.RemoveAll
		Set File = Nothing
		Form.RemoveAll
		Set Form = Nothing		
		Set coFso = Nothing
		If coUpStream.State = 1 Then coUpStream.Close
		Set coUpStream = Nothing
		If coFileStream.State = 1 Then coFileStream.Close
		Set coFileStream = Nothing		
	End Sub

'================================================================================
'== Public
'================================================================================


'--------------------------------------------------------------------------------
'-- Let Property
'--------------------------------------------------------------------------------

	' 设置文件编码
	Public Property Let CharSet(ByVal pStr)
		cvCharset = UCase(pStr)
	End Property

	' 设置单个文件最大尺寸(KB)
	Public Property Let FileMaxSize(ByVal pInt)
		cvFileMaxSize = pInt * 1024
	End Property

	' 设置所有文件最大总尺寸(KB)
	Public Property Let TotalMaxSize(ByVal pInt)
		cvTotalMaxSize = pInt * 1024
	End Property

	' 设置分块上传大小,(KB)
	Public Property Let BlockSize(ByVal pInt)
		cvBlockSize = Int(pInt) * 1024
	End Property

	' 设置允许上传的文件类型,用"|"分隔
	Public Property Let Allowed(ByVal pStr)
		cvAllowed = Replace(pStr,",","|")
	End Property

	' 设置禁止上传的文件类型,用"|"分隔
	Public Property Let Denied(ByVal pStr)
		cvDenied = Replace(pStr,",","|")
	End Property

	' 设置文件上传后保存的路径
	' 说明: 会在路径后自动加入"/"或"\"
	Public Property Let SavePath(ByVal pStr)
		' 先统一分隔符
		pStr = Replace(pStr, "\", "/")
		' 在路径后自动加入"/"或"\"
		If Instr(pStr, ":") = 2 Then
			If Right(pStr, 1) <> "/" Then pStr = pStr & "\"
		Else
			If Right(pStr, 1) <> "/" Then pStr = pStr & "/"
		End If
		cvSavePath = pStr
	End Property

	' 设置是否自动创建不存在的文件夹
	Public Property Let AutoMD(ByVal pBool)
		cvAutoMD = pBool
	End Property

	' 设置是否自动命名上传文件
	Public Property Let AutoName(ByVal pBool)
		cvAutoName = pBool
	End Property

	' 设置是否复写存在文件及文件夹
	Public Property Let Force(ByVal pBool)
		cvForce = pBool
	End Property

'--------------------------------------------------------------------------------
'-- Get Property
'--------------------------------------------------------------------------------

	' 取得单个文件最大上传限制(KB)
	Public Property Get FileMaxSize
		FileMaxSize = cvFileMaxSize / 1024
	End Property

	' 取得所有文件最大上传限制(KB)
	Public Property Get TotalMaxSize
		TotalMaxSize = cvTotalMaxSize / 1024
	End Property

	' 取得允许上传类型清单(以|分隔的字串)
	Public Property Get Allowed
		Allowed = cvAllowed
	End Property

	' 取得限制上传类型清单(以|分隔的字串)
	Public Property Get Denied
		Denied = cvDenied
	End Property

	' 取得保存文件路径
	Public Property Get SavePath
		SavePath = cvSavepath
	End Property

	' 取得错误信息
	Public Property Get ErrMsg
		ErrMsg = cvErrMsg
	End Property

'--------------------------------------------------------------------------------
'-- Function
'--------------------------------------------------------------------------------

	'-------------------------
	' 基础功能类函数
	'-------------------------

	' 创建对像操作类实例
	' 说明: 当需要创建一个新的实例时可使用此方法
	' 调用方法: Set loUpload = ebody.upload.New
	Public Function [New]()
		Set [New] = New ebody_upload
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

	' 读取并上传本地数据
	Public Sub Load
		' 基本参数
		Dim loUpFileInfo	' 记录上传文件信息的子类实例对像
		Dim lvPartData		' 记录已读取的数据
		Dim lvUploadBinary	' 记录以二进制读取上传的所有数据(用于分析上传的二进制数据)
		Dim lvSplit			' 记录上传数据每个项目之间的分隔符
		Dim lvLoadedSize	' 记录已读取到文件数据的大小(用于读取上传数据)
		Dim lvBlockSize		' 记录每次读取的数据量
		Dim lvUpDataLen		' 取得所有上传数据的大小
		Dim lvSplitLen		' 分隔符的长度
		Dim lvFormHeaderEnd	' 文件头信息的结束位置
		Dim lvCrLf			' 数据块的分隔符(回车符+换行符)
		Dim lvCrLfPre		' 预置换行分隔符
		Dim lvFormBodySize	' 文件实体数据长度
		Dim lvFormHeaderStart	' 表单文件头信息的开始位置		
		Dim lvFormHeader		' 表单文件头信息内容(二进制)
		Dim lvFormHeaderLen		' 表单文件头数据长度
		Dim lvFormBody			' 表单文件实体数据内容(二进制)
		Dim lvTotalSize			' 表单总大小		
		Dim lvFormBodyStart		' 文件实体数据开始位置
		Dim lvFormBodyEnd		' 文件实体数据结束位置
		Dim lvReadedSize		' 已经读取的数据大小(用于分解上传后的数据)
		
		' 文件信息的相关参数
		Dim lvFormHeaderStr, lvFileNameStart, lvFileNameEnd, lvFileNameLen, lvSourceFile, lvSourceName, lvSourceExt, lvSourcePath, lvSourcePathFile, lvTargetFile, lvTargetName, lvTargetPath, lvTargetPathFile, lvTargetPathUrl, lvTargetFileUrl, lvNewName
		' 表单内的项目相关参数(从文件头信息中分析读取)
		Dim lvFormNameStart, lvFormNameEnd, lvFormName, lvFormNameLen, lvFormValue
		' 图片相关参数
		Dim lgImgInfo, lvWidth, lvHeight
		
		' 数据验证检查(第一次)
		' ------------------------------------------------------------------------------

		' 检查表单是否为multipart/form-data类型
		If Not IsUploadForm_ Then cvErrMsg = "上传表单格式类型不正确" : Exit Sub

		' 检查上传表单是否有数据
		If Request.TotalBytes<1 Then cvErrMsg = "没有上传任何数据" : Exit Sub

		' 检查是否超出总尺寸限制(0表示无限制)
		If Request.TotalBytes>cvTotalMaxSize And cvTotalMaxSize>0 Then cvErrMsg = "上传数据超出总大小限制" : Exit Sub

		' 打开对像
		' ------------------------------------------------------------------------------

		' 打开接收流
		coUpStream.Type = 1	' 1 二进制模, 2 文本模式
		coUpStream.Mode = 3	' 可读可写
		coUpStream.Open		' 打开流对像

		' 打开文件流
		coFileStream.Type = 1
		coFileStream.Mode = 3
		coFileStream.Open

		' 打开fso
		Ebody.use "fso"

		' 初始已读取的文件大小
		lvLoadedSize = 0

		' 取得表单总大小		
		lvTotalSize = Request.TotalBytes

		' 分块读取上传数据(可设进度条)
		' ------------------------------------------------------------------------------

		' 循环分块读取
		Do While lvLoadedSize<lvTotalSize
			lvBlockSize = cvBlockSize
			If lvBlockSize + lvLoadedSize>lvTotalSize Then lvBlockSize = lvTotalSize - lvLoadedSize
			' 读取指定大小的内容
			lvPartData = Request.BinaryRead(lvBlockSize)
			' 统计已读大小
			lvLoadedSize = lvLoadedSize + lvBlockSize
			' 写入分块数据(在流对像后面继续添加)
			coUpStream.Write lvPartData
			' 更新进度条数据(这里可以放用于进度条的数据)
			'If b_useProgress Then o_prog.Update(lvLoadedSize)
		Loop

		' 分析上传的数据
		' ------------------------------------------------------------------------------

		' 初始上传流读取的开始位置
		coUpStream.Position = 0

		' 以二进制读取上传的所有数据(用于分析上传的二进制数据)
		lvUploadBinary = coUpStream.Read

		' 取得所有上传数据的大小
		lvUpDataLen = lvTotalSize

		' 定义数据块的分隔符(回车符+换行符)
		lvCrLf = chrB(13) & chrB(10)

		' 取得上传数据每个项目之间的分隔符(同一批提交的数据的分隔符都是一样,所以只取第一个为准)
		lvSplit = LeftB(lvUploadBinary, InStrB(1, lvUploadBinary, lvCrLf) - 1)

		' 取得分隔符的长度
		lvSplitLen = LenB(lvSplit)

		' 分解上传数据
		' ------------------------------------------------------------------------------

		Do	
			' 取得各数据位置
			' --------------------------------------------------------------------------

			' 初始值(首个上传文件的开始位置前没有lvCrLf前缀,后面的数据均有lvCrLf前缀)
			If lvFormBodyEnd>0 Then lvCrLfPre = lvCrLf

			' 取得文件头信息的开始位置(文件实体数据结束位+一个lvCrLf的长度+分隔符长度+一个lvCrLf的长度)			
			lvFormHeaderStart = lvFormBodyEnd + LenB(lvCrLfPre) + lvSplitLen + LenB(lvCrLf)

			' 取得文件头信息的结束位置(文件头与实体数据是以两个lvCrLf分隔开的)
			lvFormHeaderEnd = InStrB(lvFormHeaderStart, lvUploadBinary, lvCrLf & lvCrLf) - 1			

			' 取得文件实体数据开始位置(文件头与文件实体数据以两个lvCrLf分隔开的)
			lvFormBodyStart = lvFormHeaderEnd + LenB(lvCrLf & lvCrLf)

			' 取得文件实体数据结束位置(每个文件项是以一个lvCrLf分隔开的,从文件头的结束位开始找指定分隔符:lvCrLf+分隔符)
			lvFormBodyEnd = InStrB(lvFormHeaderEnd, lvUploadBinary, lvCrLf & lvSplit) - 1

			' 取得文件头数据长度
			lvFormHeaderLen = lvFormHeaderEnd - lvFormHeaderStart 

			' 取得文件实体数据长度
			lvFormBodySize = lvFormBodyEnd - lvFormBodyStart 

			' 取得各数据内容
			' --------------------------------------------------------------------------
	
			' 取得文件头信息内容(二进制)
			lvFormHeader = MidB(lvUploadBinary, lvFormHeaderStart, lvFormHeaderLen + 1)

			' 取得文件实体数据内容(二进制)
			lvFormBody = MidB(lvUploadBinary, lvFormBodyStart, lvFormBodySize + 1)

			' 将文件头的二进制数据转换成字符串
			lvFormHeaderStr = Bytes2Str_(lvFormHeader)

			' 取得表单内的项目名称(从文件头信息中分析读取)
			lvFormNameStart = InStr(1, lvFormHeaderStr, "name=""" , 1) + Len("name=""")
			lvFormNameEnd = InStr(lvFormNameStart, lvFormHeaderStr, """", 1)
			lvFormNameLen = lvFormNameEnd - lvFormNameStart
			lvFormName = Mid(lvFormHeaderStr, lvFormNameStart, lvFormNameLen)

			' 判断是否是文件
			If InStr(1, lvFormHeaderStr, "filename=""", 1) > 0 Then
				' 如果是文件
				' --------------------------------------------------------------------------

				' 取得文件信息
				' --------------------------------------------------------------------------

				' 取得源文件名(包含扩展名)
				lvFileNameStart = InStr(1, lvFormHeaderStr, "filename=""", 1) + Len("filename=""")
				lvFileNameEnd = InStr(lvFileNameStart, lvFormHeaderStr, """", 1)
				lvFileNameLen = lvFileNameEnd - lvFileNameStart				
				lvSourcePathFile = Replace(Mid(lvFormHeaderStr, lvFileNameStart, lvFileNameLen), "/", "\")
				lvSourceFile = Mid(lvSourcePathFile, InStrRev(lvSourcePathFile, "\") + 1)

				' 取得源文件的扩展名
				lvSourceExt = Mid(lvSourceFile, InStrRev(lvSourceFile, ".") + 1)

				' 取得源文件的名称(不含扩展名)
				' 没有扩展名,直接取文件名
				If Len(lvSourceExt) Then
					lvSourceName = Left(lvSourceFile, Len(lvSourceFile) - Len("." & lvSourceExt))
				Else
					lvSourceName = lvSourceFile
				End If

				' 取得源文件路径
				lvSourcePath = Left(lvSourcePathFile, InStrRev(lvSourcePathFile, "\"))

				' 自动编号新文件名
				If cvAutoName=True Then lvNewName = Ebody.fso.AutoName

				' 取得目标文件名(包含扩展名)
				If cvAutoName=True And Len(lvSourceExt) Then
					lvTargetFile = lvNewName & "." & lvSourceExt
				Else
					lvTargetFile = lvSourceFile
				End If

				' 取得目标文件名(不含扩展名)
				If cvAutoName=True And Len(lvTargetFile) Then
					lvTargetName = lvNewName
				Else
					lvTargetName = lvSourceName
				End If

				' 取得目标文件目录
				lvTargetPath = Ebody.MapPath(cvSavePath) & "\"

				' 取得目标文件物理地址
				lvTargetPathFile = lvTargetPath & lvTargetFile

				' 取得目标目录网址
				lvTargetPathUrl = Ebody.GetUrlAbs(cvSavePath)

				' 取得目标文件网址
				'lvTargetFileUrl = lvTargetPathUrl & Server.URLEncode(lvTargetFile)
				lvTargetFileUrl = lvTargetPathUrl & lvTargetFile
				
				' 验证数据(第二次)
				' --------------------------------------------------------------------------

				cvErrMsg = Empty	' 每个上传数据都初始错误信息
				If IsEmpty(cvErrMsg) And Not lvFormBodySize>0 Then cvErrMsg = "没有上传任何数据"	' 判断是否有上传数据(没有数据一般长度为1)
				If IsEmpty(cvErrMsg) And Not Len(lvSourceFile)>0 Then cvErrMsg = "没有上传任何数据或文件名非法"
				If IsEmpty(cvErrMsg) And Not IsAllowedType_(lvSourceExt) Then cvErrMsg = "不允许上传此类型文件"
				If IsEmpty(cvErrMsg) And cvForce=False And Ebody.fso.IsFile(lvTargetPathFile)=True Then cvErrMsg = "文件已经存在,未开启覆盖存在文件功能"
				If IsEmpty(cvErrMsg) And cvAutoMD=False And Ebody.fso.IsFolder(lvTargetPath)=False Then cvErrMsg = "上传的文件夹不存在,未开启自动创建文件夹功能"

				' 记录上传文件的成功与失败数
				If IsEmpty(cvErrMsg) Then SuccessCount = SuccessCount + 1 Else FailCount = FailCount + 1

				' 如果是图片,则取得长宽信息
				If Instr(LCase(lvFormHeaderStr),"image/") Or Instr(LCase(lvFormHeaderStr),"flash") Then
					' 初始图片长宽属性
					lvWidth = 0
					lvHeight = 0
					' 取得图片信息
					coUpStream.Position = lvFormBodyStart	' 定位数据开始位置,以可以正确读取图片信息
					lgImgInfo = GetImageSize_()	' 取得图片信息
					lvSourceExt = lgImgInfo(0)	' 扩展名
					lvWidth = lgImgInfo(1)		' 宽度
					lvHeight = lgImgInfo(2)		' 高度
					lgImgInfo = Empty
				End If

				' 建立上传文件信息类对像
				Set loUpFileInfo = New Ebody_UpFileInfo
				
				' 将文件信息分别存入文件信息类的字典中
				loUpFileInfo.SourceFile = lvSourceFile	' 源文件名(含扩展名)
				loUpFileInfo.SourceName = lvSourceName	' 源文件名(不含扩展名)
				loUpFileInfo.SourceExt = lvSourceExt	' 源文件扩展名
				loUpFileInfo.SourcePath = lvSourcePath	' 源文件路径
				loUpFileInfo.TargetFile = lvTargetFile	' 目标文件名(含扩展名)
				loUpFileInfo.TargetName = lvTargetName	' 目标文件名(不含扩展名)
				loUpFileInfo.TargetPath = lvTargetPath	' 目标文件路径
				loUpFileInfo.TargetPathFile = lvTargetPathFile	' 目标文件物理地址
				loUpFileInfo.TargetPathUrl = lvTargetPathUrl	' 目标文件所在目录的网址
				loUpFileInfo.TargetFileUrl = lvTargetFileUrl	' 目标文件网址
				loUpFileInfo.Size = FormatSize_(lvFormBodySize)	' 文件大小(已格式化)
				loUpFileInfo.FormName = lvFormName				' 表单名称
				loUpFileInfo.Width = lvWidth					' 文件宽度(如果是图片)
				loUpFileInfo.Height = lvHeight					' 文件高度(如果是图片)
				loUpFileInfo.FileStart = lvFormBodyStart		' 上传文件开始位置
				loUpFileInfo.FileSize = lvFormBodySize			' 上传文件大小(未格式化byte)
				loUpFileInfo.ErrMsg = cvErrMsg					' 错误信息
				If IsEmpty(cvErrMsg) Then loUpFileInfo.Success = True Else loUpFileInfo.Success = False	' 上传是否成功

				' 将文件对像存入文件字典(文件对像中包含有文件的各个属性,可用于前台读取)
				If NOT File.Exists(lvFormName) Then	File.Add lvFormName, loUpFileInfo
				
				' 销毁对像
				Set loUpFileInfo = Nothing

			Else	
				' 如果是表单
				' --------------------------------------------------------------------------

				' 将表单的二进制数据转换成字符串
				lvFormValue = Bytes2Str_(lvFormBody)

				' 将值存入字典
				If Form.Exists(lvFormName) Then
					' 如果存在同名对像,则累加,用,号分隔
					Form(lvFormName) = Form(lvFormName) & "," & lvFormValue
                Else
					' 新的对像,则新增新字典
					Form.Add lvFormName, lvFormValue
                End If
			End If

			' 计算出已经读取的数据大小(当前读取的数据大小+换行符长度+分隔符长度+最后一个数据的标记符长度+换行符长度)
			lvReadedSize = lvFormBodyEnd + LenB(lvCrLf) + lvSplitLen + LenB("--") + LenB(lvCrLf)

		' 判断数据是否已读完(已读数据尺寸<上传数据的总尺寸=继续读)
		Loop While lvReadedSize<lvUpDataLen

		' 关闭fso
		'Ebody.Close "fso"
	End Sub

	' 保存文件到服务器
	Public Sub Save()
		Dim lvKey, lvFileStart, lvFileSize, lvFilePath, lvFileName
		If Not IsObject(File) Then Exit Sub

		' 遍历文件字典中的所有值(文件信息对像)
		For Each lvKey In File
			' 执行文件信息对像中的保存方法
			If IsEmpty(File(lvKey).ErrMsg) Then
				' 取得文件的开始位置与文件大小等相关信息
				lvFileStart = File(lvKey).FileStart
				lvFileSize = File(lvKey).FileSize
				lvFilePath = File(lvKey).TargetPath
				lvFileName = File(lvKey).TargetFile
				' 自动创建文件夹(服务器)
				Ebody.fso.CreateFolder(lvFilePath)
				' 定位文件流的开始位置
				coUpStream.Position = lvFileStart
				' 将接收流的文件数据部份取出存入另一个目标文件流中,以用于保存成文件
				coUpStream.CopyTo coFileStream, lvFileSize
				' 将复制出来的数据流保存为文件
				coFileStream.SaveToFile lvFilePath & lvFileName, 2
			End If
		Next
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

	' 取得图片大小信息
	' 调用: GetImageSize_
	Private Function GetImageSize_()
		Dim lgImageSize(2),bFlag
		bFlag = coUpStream.Read(3)

		Select Case Hex(BinVal_(bFlag))
		Case "4E5089":
			coUpStream.Read(15)
			lgImageSize(0) = "png"
			lgImageSize(1) = BinVal2_(coUpStream.Read(2))
			coUpStream.Read(2)
			lgImageSize(2) = BinVal2_(coUpStream.Read(2))
		Case "464947":
			coUpStream.Read(3)
			lgImageSize(0) = "gif"
			lgImageSize(1) = BinVal_(coUpStream.Read(2))
			lgImageSize(2) = BinVal_(coUpStream.Read(2))
		Case "535746":
			Dim BinData,sConv,nBits
			coUpStream.Read(5)
			BinData = coUpStream.Read(1)
			sConv = Num2Str_(ASCB(BinData),2 ,8)
			nBits = Str2Num_(Left(sConv,5),2)
			sConv = Mid(sConv,6)
			While(Len(sConv)<nBits*4)
				BinData = coUpStream.Read(1)
				sConv = sConv&Num2Str_(ASCB(BinData),2 ,8)
			Wend
			lgImageSize(0) = "swf"
			lgImageSize(1) = Int(ABS(Str2Num_(Mid(sConv,1*nBits+1,nBits),2)-Str2Num_(Mid(sConv,0*nBits+1,nBits),2))/20)
			lgImageSize(2) = Int(ABS(Str2Num_(Mid(sConv,3*nBits+1,nBits),2)-Str2Num_(Mid(sConv,2*nBits+1,nBits),2))/20)
		Case "535743":'flashmx
			lgImageSize(0) = "swf"
			lgImageSize(1) = 0
			lgImageSize(2) = 0
		Case "FFD8FF":
			Dim p1
			Do
				Do: p1 = BinVal_(coUpStream.Read(1)): Loop While p1 = 255 And Not coUpStream.EOS
				If p1>191 and p1<196 Then Exit Do Else coUpStream.Read(BinVal2_(coUpStream.Read(2))-2)
				Do:p1 = BinVal_(coUpStream.Read(1)):Loop While p1<255 And Not coUpStream.EOS
			Loop While True
			coUpStream.Read(3)
			lgImageSize(0) = "jpg"
			lgImageSize(2) = BinVal2_(coUpStream.Read(2))
			lgImageSize(1) = BinVal2_(coUpStream.Read(2))
		Case Else:
			If Left(Bin2Str_(bFlag),2) = "BM" Then
				coUpStream.Read(15)
				lgImageSize(0) = "bmp"
				lgImageSize(1) = BinVal_(coUpStream.Read(4))
				lgImageSize(2) = BinVal_(coUpStream.Read(4))
			Else
				lgImageSize(0) = "(UNKNOWN)"
			End If
		End Select
		GetImageSize_ = lgImageSize
    End Function

	' 二进制转换为图片大小信息(倒取)
	Private Function BinVal_(Byval bin)
		Dim ImageSize,i
		ImageSize = 0
        For i = lenb(bin) To 1 Step -1
			ImageSize = ImageSize * 256 + ASCB(Midb(bin,i,1))
	    Next
	    BinVal_ = ImageSize
	End Function

	' 二进制转换为图片大小信息(顺取)
	Private Function BinVal2_(Byval bin)
		Dim ImageSize,i
		ImageSize = 0
		For i = 1 To Lenb(bin)
			ImageSize = ImageSize * 256 + ASCB(Midb(bin,i,1))
		Next
		BinVal2_ = ImageSize
	End Function

	' 二进制转字符
	Private Function Bin2Str_(Byval Bin)
		Dim i, Str, Sclow
		For i = 1 To LenB(Bin)
			Sclow = MidB(Bin,i,1)
			If ASCB(Sclow)<128 Then
				Str = Str & Chr(ASCB(Sclow))
			Else
				i = i+1
				If i <= LenB(Bin) Then Str = Str & Chr(ASCW(MidB(Bin,i,1)&Sclow))
			End If
		Next
		Bin2Str_ = Str
	End Function

	' 数值转字符
	Private Function Num2Str_(Byval num,Byval Base,Byval Lens)
		Dim ImageSize
		ImageSize = ""
		While(num>=Base)
				ImageSize = (num mod Base) & ImageSize
				num = (num - num mod Base)/Base
		Wend
		Num2Str_ = Right(String(Lens,"0") & num & ImageSize,Lens)
	End Function

	' 字符转数值
	Private Function Str2Num_(Byval str,Byval Base)
		Dim ImageSize,i
		ImageSize = 0
		For i=1 To Len(str)
			ImageSize = ImageSize *Base + Cint(Mid(str,i,1))
		Next
		Str2Num_ = ImageSize
	End Function

	' 把数字转换为文件大小显示方式
	Private Function FormatSize_(ByVal Size)
		If Size < 1024 Then
			FormatSize_ = FormatNumber(Size, 2) & "B"
		ElseIf Size >= 1024 And Size < 1048576 Then
			FormatSize_ = FormatNumber(Size / 1024, 2) & "KB"
		ElseIf Size >= 1048576 Then
			FormatSize_ = FormatNumber((Size / 1024) / 1024, 2) & "MB"
		End If
	End Function

	' 字节流数据转换为字符
	Private Function Bytes2Str_(ByVal pByt)
		If LenB(pByt) = 0 Then
			Bytes2Str_ = ""
			Exit Function
		End If
		Dim loStream
		Set loStream = server.createobject("ADODB.Stream")
		loStream.Type = 2	' Stream 中包含的数据的类型：1 二进制 2 文本
		loStream.Mode = 3
		loStream.Open
		loStream.WriteText pByt ' 与type对应的类型对应
		loStream.Position = 0		
		'loStream.Position = 2
		loStream.CharSet = cvCharSet
		Bytes2Str_ = loStream.ReadText() ' 返回字符串
		loStream.Close
		Set loStream = Nothing
	End Function

	'-------------------------
	' 验证类函数(Is打头)
	'-------------------------

	' 检查表单上传类型是否为:multipart/form-data
	Private Function IsUploadForm_()
		Dim lvRequestMethod
		lvRequestMethod=trim(LCase(Request.ServerVariables("REQUEST_METHOD")))
		If lvRequestMethod = "" or lvRequestMethod <> "post" Then
			IsUploadForm_ = False
			Exit Function
		End If
		Dim FormType : FormType = Split(Request.ServerVariables("HTTP_CONTENT_TYPE"), ";")
		If LCase(FormType(0)) <> "multipart/form-data" Then
			IsUploadForm_ = False
		Else
			IsUploadForm_ = True
		End If
	End Function

	' 检查上传的文件是否为允许上传类型(依扩展名来看)
	Private Function IsAllowedType_(ByVal pExt)
		' 默认全部允许
		IsAllowedType_ = True
		' 断判允许与排除的类型
		If Len(Trim(cvAllowed)) Then
			' 如果设置了仅允许上传类型,则优先
			If Not InStr(cvAllowed,pExt) > 0 Then IsAllowedType_ = False
		ElseIf Len(Trim(cvDenied)) Then			
			' 如果没有设置仅允许上传类型并设置了不允许上传类型,其次
			If InStr(cvDenied,pExt) > 0 Then IsAllowedType_ = False
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

' 内部上传文件信息子类
' 功能: 用以存储上传文件的相关信息
Class Ebody_UpFileInfo
	Public SourceFile, SourceName, SourceExt, SourcePath, TargetFile, TargetName, TargetPath, TargetPathFile, TargetPathUrl, TargetFileUrl, Size, FormName
	Public Width, Height, FileStart, FileSize, Success, ErrMsg

	Private Sub Class_Initialize()
		Width = 0
		Height = 0
		Success = False
	End Sub

	Private Sub Class_Terminate()
	End Sub
End Class
%>