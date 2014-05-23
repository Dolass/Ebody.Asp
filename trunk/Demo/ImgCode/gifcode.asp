<!--#include file="../../Ebody/ebody.asp"-->

<%
' ===============================================
' 创建基类
' ===============================================
Ebody.extend "gifcode"

' 生成gifcode实例
Dim gif : Set gif = Ebody.ext.gifcode.New

' 参数设置
gif.Noisy = 20
gif.Border = 1
gif.Angle = 1
gif.Count = Ebody.Rand(3, 8)

' step1. 一定要先创建
Session("GetCode") = gif.Create()

' step2. 再输出
gif.Output()

' 关闭
Set gif = Nothing

' ===============================================
' 销毁基类
' ===============================================
Set Ebody = Nothing
%>