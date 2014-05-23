<!--#include file="../../Ebody/ebody.asp"-->

<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" type="text/css" href="Style/EasyData/Base.css" />
<title>MD5</title>
</head>

<%
' ===============================================
' 创建基类
' ===============================================
'Dim Ebody : Set Ebody = New Ebody_Base
Ebody.CharSet = "UTF-8"


' ===============================================
' 演示 1
' 目的: 主要演示MD5功能的基本用法
' ===============================================

Response.write "<p>"
Response.write "###############################" & "<br>"
Response.write "## 主要演示MD5功能的基本用法" & "<br>"
Response.write "###############################" & "<p>"
Response.write "<p>"
Response.write "应用说明: 主要用于文件及内容的一致性效验,或数字签名,数字证书,能有效保证文件没有被更改过" & "<P>"

' 载入类
ebody.use "md5"

Dim key,Str

' 指定明文内容
str = "Abcd12345"

' 16位加密返回密文
key = ebody.md5.To16(str)
response.write "效验内容: " & str
response.write "<P>"
response.write "MD5效验字符(16位): " & key
response.write "<P>"

' 32位加密返回密文
key = ebody.md5.To32(str)
response.write "MD5效验字符(32位): " & Key

' 关闭类
ebody.close "md5"

' ===============================================
' 演示 2
' 目的: 验证文件MD5的一致性
' ===============================================

Response.write "<p>"
Response.write "###############################" & "<br>"
Response.write "## 验证文件MD5的一致性" & "<br>"
Response.write "###############################" & "<p>"
Response.write "<p>"
Response.write "应用说明: 主要用于文件及内容的一致性效验,或数字签名,数字证书,能有效保证文件没有被更改过" & "<P>"


' 载入类
ebody.use "fso"
ebody.use "md5"

Str = ebody.fso.GetFile("abc.txt")

' 16位加密返回密文
key = ebody.md5.To16(str)
response.write "效验内容: " & str
response.write "<P>"
response.write "MD5效验字符(16位): " & key
response.write "<P>"

' 32位加密返回密文
key = ebody.md5.To32(str)
response.write "MD5效验字符(32位): " & Key

' 关闭类
ebody.close "fso"
ebody.close "md5"

' ===============================================
' 销毁基类
' ===============================================
Set Ebody = Nothing

%>