<!--#include file="../../Ebody/ebody.asp"-->

<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" type="text/css" href="Style/EasyData/Base.css" />
<title>Base</title>
</head>

<%
' ===============================================
' 创建类
' ===============================================
'Dim Ebody : Set Ebody = New Ebody_Base

' ===============================================
' 演示 1
' 目的: 主要演示基本功能及属性
' ===============================================
Response.write "<p>"
Response.write "###############################" & "<br>"
Response.write "## 主要演示基本功能及属性" & "<br>"
Response.write "###############################" & "<p>"
Response.write "<p>"

Response.write "---类属性设定---" & "<p>"
Ebody.CharSet = "UTF-8"

Response.write "---类属性---" & "<p>"
Response.write "HttpType,取得服务器验证类型: " & Ebody.HttpType & "<p>"
Response.write "PortNo,取得服务器端口号: " & Ebody.PortNo & "<p>"
Response.write "Root,取得网站根物理路径: " & Ebody.Root & "<p>"
Response.write "Home,取得网站主页网址: " & Ebody.Home & "<p>"
Response.write "VRoot,自动获取当前页面所在的根物理路径: " & Ebody.VRoot & "<p>"
Response.write "VHome,自动获取当前页面所在的根网址: " & Ebody.VHome & "<p>"
Response.write "VDirName,取得当前页面所在的虚拟目录名称: " & Ebody.VDirName & "<p>"
Response.write "VDirPath,取得当前页面所在的虚拟目录的物理路径: " & Ebody.VDirPath & "<p>"
Response.write "ScriptTime,取得服务器端执行时间: "	& Ebody.ScriptTime & "<p>"
Response.write "Charset,取得页面字符集: "	& Ebody.Charset & "<p>"
Response.write "FileBOM,取得如何处理UTF-8文件的BOM信息类型: "	& Ebody.FileBOM & "<p>"
Response.write "CookieEncode,取得是否加密Cookies信息: "	& Ebody.CookieEncode & "<p>"
Response.write "Error,取得系统错误信息: "	& Ebody.Error & "<p>"

Response.write "---基础功能类函数---" & "<p>"
Response.write "IIF,判断三元表达式: " & Ebody.IIF(1=2, "YES", "NO") & "<p>"
Response.write "IsNull,判断是否为空值: " & Ebody.IsNull("") &  "<p>"
Response.write "IfThen,如果条件成立则返回某值: " & Ebody.IfThen(1=1, "成功") &  "<p>"
Response.write "IfHas,如果第1项不为空则返回第1项, 否则返回第2项: " & Ebody.IfHas("有值", "值为空") &  "<p>"
Response.write "Has,判断是否有值: " & Ebody.Has("") &  "<p>"
Response.write "Fill,不够长度的字符串填充内容: " & Ebody.Fill("abc", 10, "x") &  "<p>"
Response.write "JsEncode,处理字符串中的Javascript特殊字符: " & Ebody.JsEncode("The Path is : '\test\path'") &  "<p>"
Response.write "Test,返回正则验证结果: " & Ebody.Test("abc@taoya.com","email") &  "<p>"
Response.write "RegTest,依规则验证字符串是否存在匹配成功的结果: " & Ebody.RegTest("abc@taoya.com","^\w+([-+\.]\w+)*@(([\da-zA-Z][\da-zA-Z-]{0,61})?[\da-zA-Z]\.)+([a-zA-Z]{2,4}(?:\.[a-zA-Z]{2})?)$") &  "<p>"
Response.write "RegReplace,替换后的内容: " & Ebody.RegReplace("taoya.com", "^(([\da-zA-Z][\da-zA-Z-]{0,61})?[\da-zA-Z]\.)+([a-zA-Z]{2,4}(?:\.[a-zA-Z]{2})?)$", "domain") &  "<p>"

'Response.write "RegMatch,依规则取得结果集: "		& RegMatch("abc","b") & "<p>"
'Response.write "RegEncode,正则表达式特殊字符转义: "		& RegEncode(ByVal s) & "<p>"
'Response.write "ReplacePart,替换正则表达式编组: "		& ReplacePart(ByVal txt, ByVal rule, ByVal part, ByVal replacement) & "<p>"

Response.write "Escape,特殊字符编码: "		& Ebody.Escape("这是""Ebody""的演示文本") & "<p>"
Response.write "UnEscape,特殊字符解码: "		& Ebody.UnEscape("%u8FD9%u662F%22Ebody%22%u7684%u6F14%u793A%u6587%u672C") & "<p>"
Response.write "JsEncode,处理字符串中的Javascript特殊字符: "		& Ebody.JsEncode("abcd这是一个测试") & "<p>"
Response.write "Rand,取一个随机数: " & Ebody.Rand(3, 10) & "<p>"
Response.write "ToNumber,格式化数字: " & Ebody.ToNumber(12345.678, 2) & "<p>"
Response.write "ToPrice,将数字转换为货币格式: " & Ebody.ToPrice(233.96) & "<p>"
Response.write "ToPercent,将数字转换为百分比格式: " & Ebody.ToPercent(2.8865) & "<p>"
Response.write "CLeft,取字符隔开的左段: " & Ebody.CLeft("abc.efg", ".") & "<p>"
Response.write "CRight,取字符隔开的右段: " & Ebody.CRight("abc.eft", ".") & "<p>"
Response.write "MapPath,取得文件或文件夹在服务器上的物理存放位置(支持通配符*和?): " & Ebody.MapPath("/ebody110/dev/demo/fso/*.*") & "<p>"
Response.write "DateTime,格式化日期时间: " & Ebody.DateTime(Now, "ymmddhhiiss"&Ebody.RandStr("5:0123456789")) & "<p>"
Response.write "RandStr,取指定长度的随机字符串: " & Ebody.RandStr("5:0123456789") & "<p>"



Response.write "---设置对像类过程(Set/Remove打头)---" & "<p>"
Response.write "SetCookie,设置一个Cookies值: " & "<p>" : Call Ebody.SetCookie("EbodyCookies", "demo Cookies", 30)
'Response.write "RemoveCookie,删除一个Cookies值: " & "<p>" : Call Ebody.RemoveCookie("EbodyCookies")
Response.write "SetApp,设置缓存记录: " & "<p>" : Call Ebody.SetApp("EbodyApp", "demo app")
'Response.write "RemoveApp,删除一个缓存记录: "	 & "<p>": Call Ebody.RemoveApp("EbodyApp")

Response.write "服务器端输出javascript弹出消息" & "<p>"
'Ebody.Alert "服务器端输出javascript弹出消息"
Response.write "在服务器端输出javascript执行代码到客户端" & "<p>"
'Ebody.js "alert('删除成功！');window.parent.location.href=window.parent.location.href;"
Response.write "服务器端输出javascript弹出消息框并转到URL" & "<p>"
'Ebody.AlertUrl "服务器端输出javascript弹出消息框并转到URL", " "
Response.write "服务器端输出javascript确认消息框并根据选择转到URL" & "<p>"
'Ebody.ConfirmUrl "服务器端输出javascript确认消息框并根据选择转到URL", " ", " "

Response.write "---验证类函数(Is打头)---" & "<p>"
Response.write "IsExists,检测文件或文件夹或驱动器(磁盘)是否存在: "		& Ebody.IsExists("e:\ebody\demo.asp") & "<p>"
Response.write "IsFile,检查文件是否存在: "		& Ebody.IsFile("e:\ebody\demo.asp") & "<p>"
Response.write "IsFolder,检查目录是否存在: "		& Ebody.IsFolder("H:\WEB\Ebody110\") & "<p>"
Response.write "IsDrive,检测驱动器是否存在: "		& Ebody.IsDrive("H:") & "<p>"
Response.write "IsInstall,检测服务器组件是否安装: " & Ebody.IsInstall("Scripting.FileSystemObject") & "<p>"
Response.write "IsLoad,检测Ebody插件是否载入: " & Ebody.IsLoad("tpl") & "<p>"

Response.write "---系统取值类函数(Get打头)---" & "<p>"
Response.write "GetPathAbs,依相对路径取得文件绝对物理路径: "		& Ebody.GetPathAbs("/core/class") & "<p>"
Response.write "GetUrlAbs,依相对路径取得页面绝对网络路径: "		& Ebody.GetUrlAbs("Server/Plugin/demo.asp") & "<p>"
Response.write "GetScriptTime,脚本执行时间: " & Ebody.GetScriptTime(0) & "<p>"
Response.write "GetClientIP,获取用户IP地址: " & Ebody.GetClientIP & "<p>"
Response.write "GetCharCode,查询CodePage和Charset对照值: "		& Ebody.GetCharCode("utf-8") & "<p>"
Response.write "GetVarType,取得值类型: "		& Ebody.GetVarType("abc") & "<p>"
Response.write "GetCookie,获取一个Cookies值: "		& Ebody.GetCookie("EbodyCookies")  & "<p>"
Response.write "GetApp,获取一个缓存记录: "		& Ebody.GetApp("EbodyApp") & "<p>"

Response.write "HtmlEncode,将HTML代码转换为文本实体: "		& Ebody.HtmlEncode("<img src=abc.jpg>abc img</img>") & "<p>"
Response.write "HtmlDecode,将HTML文本转换为HTML代码: "		& Ebody.HtmlDecode("<img src=abc.jpg>abc img</img>") & "<p>"
Response.write "HtmlFilter,过滤HTML标签: "		& Ebody.HtmlFilter("<img src=abc.jpg>abc img</img>") & "<p>"




' ===============================================
' 演示 2
' 目的: 功能的应用演示
' ===============================================

Response.write "<p>"
Response.write "###############################" & "<br>"
Response.write "## 功能的应用演示" & "<br>"
Response.write "###############################" & "<p>"
Response.write "<p>"

Response.write "---tpl演示---" & "<p>"
Ebody.Use "tpl"
Ebody.Tpl.File = "tpl.txt"
Ebody.Use "fso"
Response.write Ebody.Tpl.File & "原内容: " & Ebody.FSO.GetFileAll("tpl.txt") & "<p>"
Ebody.Close "fso"
Ebody.Tpl.Tag("tag1") = "[这是include.txt中的tag1标记]"
Ebody.Tpl.Show
Response.write "<p>"

Response.write "---ext演示---" & "<p>"
'Ebody.Extend "demo"
Response.write "IsLoad ext class 1: " & Ebody.IsLoad("ext.demo") & "<p>"
'Ebody.Ext.Demo.Show
Response.write "<p>"
'Ebody.Ext.Close "demo"
Response.write "IsLoad ext class 2: " & Ebody.IsLoad("ext.demo") & "<p>"

Response.write "IsLoad ext class music: " & Ebody.IsLoad("ext.music") & "<p>"
Response.write "IsLoad ext class demo: " & Ebody.IsLoad("ext.demo") & "<p>"

Ebody.Extend "music"
Response.write "IsLoad ext class 3: " & Ebody.IsLoad("ext.music") & "<p>"
Ebody.Ext.music.Show
Response.write "<p>"
'Ebody.Ext.demo.Show
Response.write "<p>"
Ebody.Ext.Close "music"
Response.write "IsLoad ext class 4: " & Ebody.IsLoad("ext.music") & "<p>"
'Ebody.Ext.Close "demo"

Ebody.CloseExts
Response.write "IsLoad ext class 5: " & Ebody.IsLoad("ext.music") & "<p>"
Response.write "IsLoad ext class 6: " & Ebody.IsLoad("ext.demo") & "<p>"

Response.write "---类状态演示---" & "<p>"
Response.write "IsLoad tpl 1: " & Ebody.IsLoad("tpl") & "<p>"
Ebody.Close "tpl"
Response.write "IsLoad tpl 2: " & Ebody.IsLoad("tpl") & "<p>"
' ===============================================
' 注销类
' ===============================================
Set Ebody = Nothing

%>