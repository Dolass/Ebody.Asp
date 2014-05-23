<!--#include file="../../Ebody/ebody.asp"-->

<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" type="text/css" href="/Style/EasyData/Base.css" />
<title>Upload</title>
</head>


<form name="upload" method="post" action="?action=new" enctype="multipart/form-data">

请选择文件:<br>
	文字框: <input type="input" name="text" value="测试文本"><br>
	文件1: <input type="file" name="file1"><br>
	文件2: <input type="file" name="file2"><br>
	文件3: <input type="file" name="file3"><br>

	<input type="submit" value="开始上传">
</form>


<%
' ===============================================
' 创建基类
' ===============================================
'Dim Ebody : Set Ebody = New Ebody_Base
Ebody.CharSet = "UTF-8"

' ===============================================
' 演示 1
' 目的: 主要演示基本功能及属性
' ===============================================

Response.write "<p>"
Response.write "###############################" & "<br>"
Response.write "## 主要演示基本功能及属性" & "<br>"
Response.write "###############################" & "<p>"
Response.write "<p>"

' 载入类
Ebody.Use "upload"

Response.write "---类属性---" & "<p>"

Response.write "---基础功能类函数---" & "<p>"

Response.write "---验证类函数(Is打头)---" & "<p>"

Response.write "---系统取值类函数(Get打头)---" & "<p>"

' 关闭类
Ebody.Close "upload"

' ===============================================
' 演示 2
' 目的: 功能的应用演示
' ===============================================
Response.write "<p>"
Response.write "###############################" & "<br>"
Response.write "## 功能的应用演示" & "<br>"
Response.write "###############################" & "<p>"
Response.write "<p>"

' 载入类
Ebody.Use "upload"

' 下载文件
If Request("action")="download" Then
	Ebody.Use "fso"
	ebody.fso.DownloadFile("UpFiles/" & Request("file"))
	Ebody.Close "fso"
End If

' 执行上传
If Request("action")="new" Then

	' 设定属性	
	Ebody.Upload.Force = True			' 是否覆盖存在的文件
	Ebody.Upload.AutoName = True		' 是否自动命名
	Ebody.Upload.AutoMD = True			' 是否自动创建文件夹
	Ebody.Upload.CharSet = "utf-8"		' 字符编码
	Ebody.Upload.FileMaxSize = 10000	' KB
	Ebody.Upload.TotalMaxSize = 10000	' KB
	Ebody.Upload.BlockSize = 1000		' KB
	Ebody.Upload.Allowed = "jpg|jpeg|gif|doc|xls|txt|png"	' 上传文件限制
	Ebody.Upload.SavePath = "../../UpFiles"				' 相对路径,可使用..表示上级目录

	' 执行上传
	Ebody.Upload.Load

	' 保存上传的文件
	Ebody.Upload.Save

	' 显示上传信息
	Dim File, i
	For Each i In Ebody.Upload.File
		Set File = Ebody.Upload.File(i)
		'If File.FileSize > 0 Then
			' 文件信息列表
			response.write "源文件名(含扩展名): " &File.SourceFile & "<br>"
			response.write "源文件名(不含扩展名): " &File.SourceName & "<br>"
			response.write "新文件名(含扩展名): " &File.TargetFile & "<br>"
			response.write "新文件名(不含扩展名): " &File.TargetName & "<br>"
			response.write "扩展名: " &File.SourceExt & "<br>"
			response.write "原目录: " &File.SourcePath & "<br>"			
			response.write "目标路径: " &File.TargetPath & "<br>"
			response.write "目标路径网址: " &File.TargetPathUrl & "<br>"
			response.write "文件网址: " &File.TargetFileUrl & "<br>"
			response.write "文件大小(已格式化): " &File.Size & "<br>"
			response.write "表单名: " &File.FormName & "<br>"
			response.write "图片宽度: " &File.Width & "<br>"
			response.write "图片高度: " &File.Height & "<br>"
			response.write "文件开始位置: " &File.FileStart & "<br>"
			response.write "文件大小(Byte): " &File.FileSize & "<br>"
			response.write "错误信息: " &File.ErrMsg & "<br>"
			response.write "上传是否成功: " &File.Success & "<br>"
			' 下载文件
			Response.write "<a href=?action=download&file=" & File.TargetFile & ">下载文件（用于隐藏真实路径）</a>"
			response.write "<p>"
		'End If
	Next

	' 显示上传错误信息
	Response.write "成功: " & Ebody.Upload.SuccessCount & ", 失败: " & Ebody.Upload.FailCount & "<br>"
	Response.write Ebody.Upload.ErrMsg
End If

' 关闭类
Ebody.Close "upload"

' ===============================================
' 销毁基类
' ===============================================
Set Ebody = Nothing

%>