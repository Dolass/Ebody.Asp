<!--#include file="../../Ebody/ebody.asp"-->

<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" type="text/css" href="Style/EasyData/Base.css" />
<title>Ebody</title>
</head>

<%
' ===============================================
' 创建类
' ===============================================
'Dim Ebody : Set Ebody = New Ebody_Base

' 载入类
Ebody.extend "music"

ebody.ext.music.show

' 关闭类
Ebody.Close "music"

' ===============================================
' 注销类
' ===============================================
Set Ebody = Nothing

%>