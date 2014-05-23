<!--#include file="../../Ebody/ebody.asp"-->

<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" type="text/css" href="Style/EasyData/Base.css" />
<title>TPL</title>
</head>

<%
' ===============================================
' 创建基类
' ===============================================
'Dim Ebody : Set Ebody = New Ebody_Base
Ebody.CharSet = "UTF-8"

' ===============================================
' 演示 1
' 目的: 主要演示自动更新块内容的使用方法
' ===============================================

Response.write "<p>"
Response.write "###############################" & "<br>"
Response.write "## 主要演示自动更新块内容的使用方法" & "<br>"
Response.write "###############################" & "<p>"
Response.write "<p>"

' 载入类
Ebody.Use "tpl"
' 设置未定义标签的处理方式
Ebody.Tpl.UnDefine = "keep"
' 载入模板文件
Ebody.Tpl.File = "model/tpl.htm"

' 绑定数据
ebody.tpl.tag("B.title") = "B_title"
ebody.tpl.tag("A.title") = "A_title"
ebody.tpl.tag("A.addtime") = "A_addtime"
ebody.tpl.tag("id") = "Tag_id"
'ebody.tpl.tag("A") = "Block_A"
'ebody.tpl.tag("B1") = "Block_B1"
ebody.tpl.tag("B2") = "Block_B2"

' 自动显示填充后的结果
Ebody.Tpl.Show

' 关闭类
Ebody.Close "tpl"

' ===============================================
' 演示 2
' 目的: 主要演示手动更新块内容的使用方法
' ===============================================

Response.write "<p>"
Response.write "###############################" & "<br>"
Response.write "## 主要演示手动更新块内容的使用方法" & "<br>"
Response.write "###############################" & "<p>"
Response.write "<p>"

' 载入类
Ebody.Use "tpl"
' 设置未定义标签的处理方式
Ebody.Tpl.UnDefine = "Remove"
' 载入模板文件
Ebody.Tpl.File = "Model/tpl.htm"

' 绑定数据
ebody.tpl.tag("B.title") = "B_title"
ebody.tpl.tag("A.title") = "A_title"
ebody.tpl.tag("A.addtime") = "A_addtime"
ebody.tpl.tag("id") = "Tag_id"
ebody.tpl.tag("A") = "Block_A"

' 手动更新块内容
Ebody.Tpl.UpdateBlock "A"
Ebody.Tpl.UpdateBlock "B2"

' 手动构建页面内容
Ebody.Tpl.Build

' 显示构建后的页面内容
Response.write Ebody.Tpl.html

' 关闭类
Ebody.Close "tpl"

' ===============================================
' 演示 3
' 目的: 主要演示将构建后的页面存为文件
' ===============================================

Response.write "<p>"
Response.write "###############################" & "<br>"
Response.write "## 主要演示将构建后的页面存为文件" & "<br>"
Response.write "###############################" & "<p>"
Response.write "<p>"

' 载入类
Ebody.Use "tpl"
Ebody.Use "fso"

' 设置未定义标签的处理方式
Ebody.Tpl.UnDefine = "keep"
' 载入模板文件
Ebody.Tpl.File = "model/tpl.htm"

' 绑定数据
ebody.tpl.tag("B.title") = "B_title"
ebody.tpl.tag("A.title") = "A_title"
ebody.tpl.tag("A.addtime") = "A_addtime"
ebody.tpl.tag("id") = "Tag_id"
ebody.tpl.tag("A") = "Block_A"

' 手动更新块内容
Ebody.Tpl.UpdateBlock "A"
Ebody.Tpl.UpdateBlock "B2"

' 手动构建页面内容
Ebody.Tpl.Build

' 内容存为文件
If Not Ebody.fso.isFile("Files/index.txt") Then 
	Ebody.fso.CreateFile "Files/index.txt", Ebody.tpl.HTML
End If

' 文件生成地址
Response.write "主机网址：" & Ebody.VHome & "<P>"
Response.write "文件网址：" & Ebody.GetUrlAbs("Files/index.txt") & "<P>"
Response.write "文件路径：" & Ebody.GetPathAbs("Files/index.txt") & "<P>"

' 查看生成的文件
Response.write "<a href=" & Ebody.GetUrlAbs("Files/index.txt") & " target=_blank>查看生成的文件</a>" & "<P>"

' 下载生成的文件
Response.write "<a href= ?action=download>下载生成的文件</a>" & "<P>"
If Request("action")="download" Then
	ebody.fso.DownloadFile("Files/index.txt")
End If

' 关闭类
Ebody.Close "fso"
Ebody.Close "tpl"

' ===============================================
' 演示 4
' 目的: tpl一些属性的演示
' ===============================================

Response.write "<p>"
Response.write "###############################" & "<br>"
Response.write "## tpl一些属性的演示" & "<br>"
Response.write "###############################" & "<p>"
Response.write "<p>"

' 载入类
Ebody.Use "tpl"
' 设置未定义标签的处理方式
Ebody.Tpl.UnDefine = "keep"
' 载入模板文件
Ebody.Tpl.Str = "<!-- {#:A} -->[{A.title}][{A.addtime}][{id}][{B.title}]<!-- {#:B} --><!-- {#:C} -->[{A.title}]<!-- {/#:C} -->[{B.title}]<!-- {/#:B} --><!-- {/#:A} -->"

' 绑定数据
ebody.tpl.tag("B.title") = "B_title"
ebody.tpl.tag("A.title") = "A_title"
ebody.tpl.tag("A.addtime") = "A_addtime"
ebody.tpl.tag("id") = "Tag_id"
'ebody.tpl.tag("A") = "Block_A"
'ebody.tpl.tag("B") = "Block_B"
'ebody.tpl.tag("C") = "Block_C"

' 手动更新块内容
Ebody.Tpl.UpdateBlock "A"
Ebody.Tpl.UpdateBlock "B"
Ebody.Tpl.UpdateBlock "C"

' 显示属性值
Response.write "TagData(A.title): " & Ebody.tpl.TagData("A.title") & "<P>"
Response.write "TagData(B.title): " & Ebody.tpl.TagData("B.title") & "<P>"
Response.write "TagData(id): " & Ebody.tpl.TagData("id") & "<P>"
Response.write "Block(A): " & Ebody.tpl.Block("A") & "<P>"
Response.write "BlockData(A): " & Ebody.tpl.BlockData("A") & "<P>"
Response.write "Block(B): " & Ebody.tpl.Block("B") & "<P>"
Response.write "BlockData(B): " & Ebody.tpl.BlockData("B") & "<P>"
Response.write "BlockDataAll(B): " & Ebody.tpl.BlockDataAll("B") & "<P>"

' 手动构建页面内容
Ebody.Tpl.Build

' 显示构建后的页面内容
Response.write "html: " & Ebody.Tpl.html

' 关闭类
Ebody.Close "tpl"

' ===============================================
' 演示 5
' 目的: 多语言支持
' ===============================================

Response.write "<p>"
Response.write "###############################" & "<br>"
Response.write "## 多语言支持" & "<br>"
Response.write "###############################" & "<p>"
Response.write "<p>"

' 多语言支持
Ebody.Include "../lang/" & Ebody.IfHas(request("lang"),"CN") & ".asp" ' 加载对应的语言包
Response.write lang(5) & ":" & "<P>"
Response.write "<a href=?lang=en>English</a>" & "<P>"
Response.write "<a href=?lang=cn>中文</a>" & "<P>"
Response.write "<P>"
Response.write lang(1)


' ===============================================
' 销毁基类
' ===============================================
Set Ebody = Nothing

%>