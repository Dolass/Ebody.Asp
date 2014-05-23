<!--#include file="../../Ebody/ebody.asp"-->

<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" type="text/css" href="/Style/EasyData/Base.css" />
<title>FSO</title>
</head>

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
Ebody.Use "fso"

Response.write "---类属性---" & "<p>"
Response.write "ThisFilePath,取得当前页面的物理文件地址: " & ebody.fso.ThisFilePath & "<p>"
Response.write "ThisDirPath,取得当前页面的物理目录地址: " & ebody.fso.ThisDirPath & "<p>"
Response.write "RPath,取得当前页面基于网站主页的相对网络目录路径: " & ebody.fso.RPath & "<p>"
Response.write "ThisFile,取得当前页面的文件名(包含扩展名): " & ebody.fso.ThisFile & "<p>"
Response.write "ThisFileName,取得当前文件名(不含扩展名): " & ebody.fso.ThisFileName & "<p>"
Response.write "ThisFileExt,取得当前文件的扩展名: " & ebody.fso.ThisFileExt & "<p>"

Response.write "---基础功能类函数---" & "<p>"
Response.write "AutoName,自动命名: " & ebody.fso.AutoName & "<p>"
Response.write "CreateFolder,创建文件夹: " & ebody.fso.CreateFolder("Folder1/Folder2") & "<p>"
Response.write "CreateFile,将文本数据保存为文件: " & ebody.fso.CreateFile("Folder1/Folder2/CreateFile.txt", "abc/") & "<p>"
'Response.write "将二进制数据保存为文件：" & ebody.fso.SaveAs("Folder1/Folder2/CreateFile.txt", "efg/") & "<p>"
Response.write "CreateTextFile,向文本文件写入内容（覆写）: " & ebody.fso.CreateTextFile("Folder1/Folder2/CreateFile.txt", "efg/") & "<p>"
Response.write "AppendFile,向文本文件追加内容: " & ebody.fso.AppendTextFile("Folder1/Folder2/CreateFile.txt", "efg/") & "<p>"
Response.write "Rename,重命名文件或文件夹: " & ebody.fso.Rename("Folder1/Folder2/CreateFile.txt", "NewName.txt") & "<p>"
Response.write "CopyFolder,复制文件夹: " & ebody.fso.CopyFolder("Folder1/Folder2", "Folder1/Folder3") & "<p>"
Response.write "MoveFile,移动文件: " & ebody.fso.MoveFile("Folder1/Folder2/*.*", "Folder1/Folder3") & "<p>"
Response.write "DelFolder,删除文件夹: " & ebody.fso.DelFolder("Folder1/Folder2") & "<p>"
Response.write "DelFile,删除文件: " & ebody.fso.DelFile("Folder1/Folder3/New*.txt") & "<p>"

Response.write "---验证类函数(Is打头)---" & "<p>"
Response.write "IsExists,检测文件或文件夹或驱动器(磁盘)是否存在: "		& Ebody.IsExists("e:\ebody\demo.asp") & "<p>"
Response.write "IsFile,检查文件是否存在: "		& Ebody.IsFile("e:\ebody\demo.asp") & "<p>"
Response.write "IsFolder,检查目录是否存在: "		& Ebody.IsFolder("Folder1/Folder3") & "<p>"
Response.write "IsDrive,检测驱动器是否存在: "		& Ebody.IsDrive("H:") & "<p>"

Response.write "---系统取值类函数(Get打头)---" & "<p>"
Response.write "GetFileName,取文件名: " & ebody.fso.GetFileName("common/abc/abc.asp") & "<p>"
Response.write "GetFileExt,取文件扩展名: " & ebody.fso.GetFileExt("common/abc/abc.asp") & "<p>"
Response.write "GetFile,读取指定文本文件内容: " & ebody.fso.GetFile("Folder1/Folder3/Folder2/NewName.txt") & "<p>"
Response.write "GetFileAll,无限级读取文件内容, 同时也读取所有包含文件的内容: " & ebody.fso.GetFileAll("Folder1/Folder3/Folder2/NewName.txt") & "<p>"

Response.write "---文件/文件夹属性---<p>"
Response.write "资源名称: " & ebody.fso.GetAttr("Folder1/Folder3/Folder2/NewName.txt", "name") & "<p>"
Response.write "上次修改时间: " & ebody.fso.GetAttr("Folder1/Folder3/Folder2/NewName.txt", "date") & "<p>"
Response.write "创建时间: " & ebody.fso.GetAttr("Folder1/Folder3/Folder2/NewName.txt", "datecreated") & "<p>"
Response.write "上次访问时间: " & ebody.fso.GetAttr("Folder1/Folder3/Folder2/NewName.txt", "dateaccessed") & "<p>"
Response.write "资源大小: " & ebody.fso.GetAttr("Folder1/Folder3/Folder2/NewName.txt", "size") & "<p>"
Response.write "资源属性: " & ebody.fso.GetAttr("Folder1/Folder3/Folder2/NewName.txt", "attr") & "<p>"
Response.write "资源类型: " & ebody.fso.GetAttr("Folder1/Folder3/Folder2/NewName.txt", "type") & "<p>"
Response.write "资源物理路径: " & ebody.fso.GetAttr("Folder1/Folder3/Folder2/NewName.txt", "path") & "<p>"
Response.write "父文件夹: " & ebody.fso.GetAttr("Folder1/Folder3/Folder2/NewName.txt", "parentfolder") & "<p>"

Response.write "---驱动器属性---<p>"
Response.write "根文件夹: " & ebody.fso.GetAttr("d:", "rootfolder") & "<p>"
Response.write "设置或返回指定驱动器的卷标名: " & ebody.fso.GetAttr("d:", "volumename") & "<p>"
Response.write "驱动器或网络共享的总空间大小: " & ebody.fso.GetAttr("d:", "totalsize") & "<p>"
Response.write "指定驱动器上或共享驱动器可用的磁盘空间: " & ebody.fso.GetAttr("d:", "freespace") & "<p>"
Response.write "指定的驱动器或网络共享上的用户可用的空间容量: " & ebody.fso.GetAttr("d:", "availablespace") & "<p>"
Response.write "取得磁盘格式类型(可用的返回类型包括 FAT、NTFS 和 CDFS): " & ebody.fso.GetAttr("d:", "filesystem") & "<p>"

Response.write "---基础功能类过程---<p>"
'Response.write "DownloadFile,自动下载文件: " & ebody.fso.DownloadFile("Folder1/NewName.txt") & "<p>"
Response.write "根文件夹: " & ebody.fso.GetAttr("d:", "rootfolder") & "<p>"

' 关闭类
Ebody.Close "fso"

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
Ebody.Use "fso"

Response.write "---设定属性---" & "<p>"
Response.write "设定文件隐藏/只读: " & ebody.fso.SetAttr("Folder1/Folder3/Folder2/NewName.txt", "+H,+R") & "<p>"

' 关闭类
Ebody.Close "fso"

' ===============================================
' 销毁基类
' ===============================================
Set Ebody = Nothing

%>