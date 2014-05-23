<!--#include file="../../Ebody/ebody.asp"-->

<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" type="text/css" href="Style/EasyData/Base.css" />
<title>AES</title>
</head>

<%
' ===============================================
' 创建基类
' ===============================================
'Dim Ebody : Set Ebody = New Ebody_Base
Ebody.CharSet = "UTF-8"


' ===============================================
' 演示 1
' 目的: 主要演示AES功能的基本用法
' ===============================================

Response.write "<p>"
Response.write "###############################" & "<br>"
Response.write "## 主要演示AES功能的基本用法" & "<br>"
Response.write "###############################" & "<p>"
Response.write "<p>"
Response.write "应用说明: 主要用于密码字符加密" & "<P>"

' 载入类
ebody.use "aes"

Dim str,key,show

' 指定明文内容
str = "12345abcdE"

' 设置和读取AES密钥
ebody.aes.Password = "abc"
Response.write "密钥：" & ebody.aes.Password & "<P>"

' AES加密
key = ebody.aes.Encode(str)
response.write "密码经AES加密字符: " & Key & "<P>"

' AES解密
show = ebody.aes.Decode(key)
response.write "原密码字符: " & show

' 关闭类
ebody.close "aes"

' ===============================================
' 销毁基类
' ===============================================
Set Ebody = Nothing

%>