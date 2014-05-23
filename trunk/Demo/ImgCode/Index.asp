<!--#include file="../../Ebody/ebody.asp"-->

<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" type="text/css" href="Style/EasyData/Base.css" />
<title>IMG CODE</title>
</head>

<%

Response.write "<a href=?type=gif>GIF图</a>" & "<P>"
Response.write "<a href=?type=bmp>BMP图</a>" & "<P>"

' ===============================================
' 创建基类
' ===============================================
'Dim Ebody : Set Ebody = New Ebody_Base

' 输出方法一
If request("type") = "gif" Then	
	Response.write "<img src=code_gif.asp>" & "<br>"
	Response.write "验证内容：" & Session("GetCode") & "<br>"
	
	Response.write "<form name=validate action=?action=login method=post>"
	Response.write "请输入验证码：<input type=text name=code>"
	Response.write "<input type=submit value=确定></form>" & "<br>"
Else
	Response.write "<img src=code_bmp.asp>" & "<br>"
	Response.write "验证内容：" & Session("CheckCode") & "<br>"
End If

' 输出方法二
Response.write "<img src=gifcode.asp>"


' ===============================================
' 销毁基类
' ===============================================
Set Ebody = Nothing

' 验证代码
If request("action") = "login" Then
	If Session("GetCode") = request("code") Then
		Response.write "你的验证成功！"
	Else
		Response.write "验证失败！"
	End If
End If

%>