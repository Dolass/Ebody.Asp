<!--#include file="../../Ebody/ebody.asp"-->

<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" type="text/css" href="Style/EasyData/Base.css" />
<title>Mult Lang</title>
</head>

<%
' ===============================================
' 创建基类
' ===============================================
'Dim Ebody : Set Ebody = New Ebody_Base
Ebody.CharSet = "UTF-8"

' ===============================================
' 演示 1
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

Response.write "<P>"
''Response.write "Request.ServerVariables:" & Request.ServerVariables("Server_Name")
'Response.write VHome & Replace("../../UpFiles", "\", "/")
Dim lvPath : lvPath = Request.ServerVariables("Script_Name")
				GetUrlAbs = Home & Left(lvPath, InStrRev(lvPath, "/") - 1) & "/" & "../../UpFiles"
				Response.write GetUrlAbs
Response.write "<P>"
Response.write Request.ServerVariables("SCRIPT_NAME")

' ===============================================
' 销毁基类
' ===============================================
Set Ebody = Nothing

%>