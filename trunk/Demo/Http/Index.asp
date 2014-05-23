<!--#include file="../../Ebody/ebody.asp"-->

<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" type="text/css" href="Style/EasyData/Base.css" />
<title>HTTP</title>
</head>

<%
' ===============================================
' 创建基类
' ===============================================
'Dim Ebody : Set Ebody = New Ebody_Base
Ebody.CharSet = "UTF-8"
Dim http, tmp, rule, arr, i

' ====================================================================================================

'Response.write "<p>"
'Response.write "###############################" & "<br>"
'Response.write "## Demo 1 - 最简单的应用(Get)" & "<br>"
'Response.write "###############################" & "<p>"
'Response.write "<p>"
'
'
' 载入类
Ebody.Use "http"

' 直接获取页面源码（小偷程序）
tmp = ebody.Http.Get("http://www.baidu.com")
Response.write "目标HTML代码：" & ebody.HtmlEncode(tmp)

' 生成文件
ebody.use "fso"
ebody.fso.createfile "baidu.html", tmp
ebody.close "fso"

' 关闭类
ebody.close "http"

' ====================================================================================================

'Response.write "<p>"
'Response.write "###############################" & "<br>"
'Response.write "## Demo 2 - 最简单的Post" & "<br>"
'Response.write "###############################" & "<p>"
'Response.write "<p>"
'
'' 载入类
'Ebody.Use "http"
'
'' 取得带参数查询返回的结果*****
'ebody.Http.Data = Array("srchtxt:钛架","srchmod:tech")
'tmp = ebody.Http.Post("http://s.weiphone.com/search.php")
'Response.write ebody.HtmlEncode(tmp)
'
'' 关闭类
'ebody.close "http"

' ====================================================================================================

'Response.write "<p>"
'Response.write "###############################" & "<br>"
'Response.write "## Demo 3 - 通过属性配置" & "<br>"
'Response.write "###############################" & "<p>"
'Response.write "<p>"
'
'' 载入类
'Ebody.Use "http"
'
'Set http = ebody.Http.New()
'http.ResolveTimeout = 20000	'服务器解析超时时间，毫秒，默认20秒
'http.ConnectTimeout = 20000	'服务器连接超时时间，毫秒，默认20秒
'http.SendTimeout = 300000		'发送数据超时时间，毫秒，默认5分钟
'http.ReceiveTimeout = 60000	'接受数据超时时间，毫秒，默认1分钟
'http.Url = "http://s.weiphone.com/search.php"	'目标URL地址
'http.Method = "POST"  'GET 或者 POST, 默认GET
'
''目标文件编码，一般不用设置此属性，ebody会自动判断目标地址的编码
''http.CharSet = "gb2312"
'http.Async = False	'异步，默认False，建议不要修改
'
''数据提交方式一，如果是GET则会附在URL后以参数形式提交：
''http.Data = "srchmod=tech&srchtxt=" & Server.URLEncode("钛架")
'
''数据提交方式二，可以用Array参数的方式提交：
'http.Data = Array("srchtxt:钛架","srchmod:tech")
'http.User = ""	'如果访问目标URL需要用户名
'http.Password = ""	'如果访问目标URL需要密码
'
'' 打开请求，并返回结果
'http.Open
'Response.write ebody.HtmlEncode(http.Html)
'Set http = Nothing
'
'' 关闭类
'ebody.close "http"

' ====================================================================================================

'Response.write "<p>"
'Response.write "###############################" & "<br>"
'Response.write "## Demo 4 - 获取文件头" & "<br>"
'Response.write "###############################" & "<p>"
'Response.write "<p>"
'
'' 载入类
'Ebody.Use "http"
'
'ebody.Http.Get "http://www.baidu.com"
'tmp = ebody.Http.Headers
'Response.write ebody.HtmlEncode(tmp)
'
'' 关闭类
'ebody.close "http"

' ====================================================================================================

'Response.write "<p>"
'Response.write "###############################" & "<br>"
'Response.write "## Demo 5 - 获取文件指定部分内容" & "<br>"
'Response.write "###############################" & "<p>"
'Response.write "<p>"
'
'' 载入类
'Ebody.Use "http"
'
'Dim bookid,bookname,bookdesc,uptime,readlink
'
'bookid = "thread-830155-1-1"
'tmp = Ebody.Http.Get("http://bbs.biketo.com/"&bookid&".html")
'
''Response.write ebody.HtmlEncode(tmp)
'
'' 用SubStr按字符截取部分文本
'bookname = ebody.Http.SubStr("<a href=""thread-830155-1-1.html"" id=""thread_subject"">","</a>",0)
'bookdesc = ebody.Http.SubStr("<td class=""t_f"" id=""postmessage_4806729"">","</td>",0)
'
'' 用Find可按正则获取一段文本
'uptime = ebody.Http.Find("发表于 [\d- :]+")
'
'' 用Select可按正则编组选择匹配的部分文本,$0是获取正则匹配的字符串本身
'readlink = ebody.Http.Select("(<a href="")(/thread/\d+.html)(.+</a>)","$1http://bbs.biketo.com$2$3")
'
'Response.write "<b>标题：</b>" & bookname & "<P>"
'Response.write "<b>发表于: </b>" & uptime & "<P>"
'Response.write "<b>阅读地址：</b>" & readlink & "<P>"
'Response.write "<b>内容简介：</b>" & bookdesc & "<P>"
'
'' 关闭类
'ebody.close "http"

' ====================================================================================================

Response.write "<p>"
Response.write "###############################" & "<br>"
Response.write "## Demo 6 - 获取文件循环部分" & "<br>"
Response.write "###############################" & "<p>"
Response.write "<p>"

' 载入类
Ebody.Use "http"

ebody.Http.Get "http://code.google.com/p/easyasp/updates/list"
rule = "<span class=""date below-more"" title=""(.+?)""[\s\S]+?>(.+?)</span>[\s\S]+?<span class=""title""><a class=""ot-revision-link"" href=""/p/easyasp/source/detail\?r=(?:\d+?)"">(r\d+?)</a>\n \(([\s\S]+?)\).+>(\w+?)</a></span>"
arr = ebody.Http.Search(rule)
'Easp.WN "====前5个匹配===="
'For i = 0 To 4
'	Easp.WN "<b>第" & i + 1 & "个匹配项：</b>"
'	Easp.WN Easp.HtmlEncode(arr(i))
'Next
'Easp.WN ""
'还可以用正则来进行更复杂的应用
Dim Matches, Match
Set Matches = ebody.RegMatch(ebody.Http.Html,rule)
Response.write "====EasyASP更新日志摘要===="
For Each Match In Matches
	If Match.SubMatches(3)<>"[No log message]" Then Response.write ebody.Format("<li>{3}, {4} ({5} @ {2})</li>",Match)
Next
Set Matches = Nothing

' 关闭类
ebody.close "http"

' ====================================================================================================

'Response.write "<p>"
'Response.write "###############################" & "<br>"
'Response.write "## Demo 7 - 保存远程图片" & "<br>"
'Response.write "###############################" & "<p>"
'Response.write "<p>"
'
'' 载入类
'Ebody.Use "http"
'
'ebody.Http.Get "http://www.baidu.com"
'tmp = ebody.Http.SaveImgTo("imgatlocal/")
'Response.write ebody.HtmlEncode(tmp)
'
'' 关闭类
'ebody.close "http"

' ====================================================================================================

'Response.write "<p>"
'Response.write "###############################" & "<br>"
'Response.write "## Demo 8 - WebService(SOAP1.1)示例，获得腾讯QQ在线状态 <br>"
'Response.write "###############################" & "<p>"
'Response.write "<p>"
'
'' 载入类
'Ebody.Use "http"
'
'''获得腾讯QQ在线状态
'Dim QQ,xml : QQ = 46220480
'tmp = "<?xml version=""1.0"" encoding=""utf-8""?>"
'tmp = tmp & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
'tmp = tmp & "  <soap:Body>"
'tmp = tmp & "    <qqCheckOnline xmlns=""http://WebXml.com.cn/"">"
'tmp = tmp & "      <qqCode>" & QQ & "</qqCode>"
'tmp = tmp & "    </qqCheckOnline>"
'tmp = tmp & "  </soap:Body>"
'tmp = tmp & "</soap:Envelope>"
'Set http = ebody.Http.New
'
''设置请求头信息的三种方式
''其一：
''http.RequestHeader("Host") = "www.webxml.com.cn"
''其二：
'http.SetHeader "Content-Type:text/xml; charset=utf-8"
''其三：
''http.SetHeader Array("Content-Length:" & Len(tmp), "SOAPAction:http://WebXml.com.cn/qqCheckOnline")
'
'' 取得响应
'http.Data = tmp
'tmp = http.Post("http://www.webxml.com.cn/webservices/qqOnlineWebService.asmx?WSDL")
'Set http = Nothing
'
''解析返回数据
'Ebody.Use "xml"
'Set xml = Ebody.Xml.New
'xml.Load tmp
'tmp = xml("qqCheckOnlineResult").Value
'Set xml = Nothing
'Select Case tmp
'	Case "Y" tmp = "在线"
'	Case "N" tmp = "离线"
'	Case "E" tmp = "号码错误"
'	Case "A" tmp = "商业用户验证失败"
'	CAse "V" tmp = "免费用户超过数量"
'End Select
'Response.write "QQ:" & QQ & " (" & tmp & ")"
'
'' 关闭类
'ebody.close "xml"
'ebody.close "http"


'' 设置代理，通过其它html页面调用有此设定的页面，来返回实时数据，类同ajax
'ebody.Use "Http"
'Easp.Http.AjaxAgent()
'ebody.close "http"

'=========================

Response.write "<P>------------------------------------<P>"
Response.write "页面执行时间： " & Ebody.ScriptTime & " 秒"

%>