<!--#include file="../../Ebody/ebody.asp"-->

<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" type="text/css" href="Style/EasyData/Base.css" />
<title>综合DEMO</title>
</head>

<%
' ===============================================
' 创建基类
' ===============================================
'Dim Ebody : Set Ebody = New Ebody_Base
Ebody.CharSet = "UTF-8"

' 载入类
Ebody.Use "http"

' 直接获取页面源码（小偷程序）
tmp = ebody.Http.Get("http://www.sina.com")
'Response.write ebody.HtmlEncode(tmp)
Response.write tmp

' 关闭类
ebody.close "http"

%>