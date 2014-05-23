<!--#include file="../../Ebody/ebody.asp"-->

<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" type="text/css" href="Style/EasyData/Base.css" />
<title>XML</title>
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

'Response.write "<p>"
'Response.write "###############################" & "<br>"
'Response.write "## 主要演示XML功能的基本用法" & "<br>"
'Response.write "###############################" & "<p>"
'Response.write "<p>"

' ====================================================================================================

Response.write "<p>"
Response.write "###############################" & "<br>"
Response.write "## 读取xml数据" & "<br>"
Response.write "###############################" & "<p>"
Response.write "<p>"

' 载入类
Ebody.Use "xml"

Dim str,n,i

str = 			"<?xml version=""1.0"" encoding=""utf-8""?>" & vbCrLf
str = str & "<microblog>" & vbCrLf
str = str & "	<site>" & vbCrLf
str = str & "		<name alias=""Tencent"">腾讯微博</name>" & vbCrLf
str = str & "		<url>http://t.qq.com</url>" & vbCrLf
str = str & "		<account nick=""user"" for=""me""><name>@lengshi</name><nick>Ray</nick></account>" & vbCrLf
str = str & "		<last><![CDATA[今天我们这里下<em>大雨</em>啦！]]></last></site>" & vbCrLf
str = str & "	<site>" & vbCrLf
str = str & "		<name alias=""Sina"">新浪微博</name>" & vbCrLf
str = str & "		<url>http://t.sina.com.cn</url>" & vbCrLf
str = str & "		<account nick=""email"" for=""me""><name>@tainray</name><nick>tainray@sina.com</nick></account>" & vbCrLf
str = str & "		<last><![CDATA[是不是<font color=""red"">这样</font>的噢，我也不知道哈。<img src=""http://bbs.lengshi.com/max-assets/icon-emoticon/12.gif"" />]]></last></site>" & vbCrLf
str = str & "	<site>" & vbCrLf
str = str & "		<name alias=""Twitter"">推特</name>" & vbCrLf
str = str & "		<url>http://twitter.com</url>" & vbCrLf
str = str & "		<account nick=""user"" for=""notme""><name haha=""1"">@ccav</name><nick>CCAV</nick></account>" & vbCrLf
str = str & "		<last><![CDATA[I don't need this feature <strong>(>_<)</strong> any more.]]></last></site>" & vbCrLf
str = str & "</microblog>"

'载入Xml数据
'Ebody.Xml.Load "http://Ebody.lengshi.cn/data/xml/microblog_catalog.xml"	' 载入远程数据
'Ebody.Xml.Open "microblog.xml"		' 载入本地数据
Ebody.Xml.Load str	' 载入字符内容
'选择所有标签为name的节点，并输出找到的节点个数
Response.write "选择所有标签为name的节点，并输出找到的节点个数："
Response.write Ebody.Xml("name").Length & "<P>"
Response.write "--------" & "<P>"
'选择所有包含属性alias的标签为name的节点
Response.write "选择所有包含属性alias的标签为name的节点："
Response.write Ebody.Xml("name[alias]").Length & "<P>"
Response.write "--------" & "<P>"
'选择所有属性for等于me，nick属性不等于email的标签为account的节点，并输出其Xml代码
Response.write "选择所有属性for等于me，nick属性不等于email的标签为account的节点，并输出其Xml代码："
Response.write Ebody.Xml("account[for='me'][nick!='email']").Xml & "<P>"
Response.write "--------" & "<P>"
'选择site节点的子节点中标签为name的节点
Response.write "选择site节点的子节点中标签为name的节点："
Response.write Ebody.Xml("site>name").Xml & "<P>"
Response.write "--------" & "<P>"
'选择account节点的后代节点中标签为name的节点
Response.write "选择account节点的后代节点中标签为name的节点："
Response.write Ebody.Xml("account name").Xml & "<P>"
Response.write "--------" & "<P>"
'选择所有的url和last节点
Response.write "选择所有的url和last节点：" & "<P>"
Response.write Ebody.HtmlEncode(Ebody.Xml("url,last").Xml) & "<P>"
Response.write "--------" & "<P>"
'修改指定节点内容
Response.write "修改指定节点内容：" & "<P>"
Response.write "修改前：" & Ebody.HtmlEncode(Ebody.Xml("url")(0).Xml) & "<P>"
Ebody.Xml("url")(0).Text = "<test>sss</test>"
Response.write "修改后：" & Ebody.HtmlEncode(Ebody.Xml("url")(0).Xml) & "<P>"


'Response.write "--------" & "<P>"
'Response.write "以xslt模式读取Xml内容" & "<P>"
'Ebody.Xml.XSLT = "xsl/microblog.xsl"
'Response.write Ebody.HtmlEncode(Ebody.Xml.Dom.Xml)
'
'Response.write "--------" & "<P>"
'Response.write "将内容保存为xml文件" & "<P>"
'Response.write Ebody.Xml.SaveAs("news.xml>gbk")
'Response.write Ebody.Xml.SaveAs("microblog.xml>utf-8")

'Response.write "--------" & "<P>"
'Response.write "依标记取得标记内的值：" & "<P>"
'
'Set n = Ebody.Xml("url")
'For i = 0 To n.Length-1
'	Response.write n(i).Value & "<P>"
'Next
'Set n = Nothing
'
'Response.write "--------" & "<P>"
'Response.write "依标记取得标记内的值：" & "<P>"
'Response.write Ebody.Xml("last")(2).Value & "<P>"
'Set n = Ebody.Xml("last")
'For i = 0 To n.Length-1
'	'Ebody.WN n(i).Type
'	Response.write n(i).Value & "<P>"
'Next
'Response.write n.Text & "<P>"
'Response.write n(1).Root.Type & "<P>"
'Response.write n(2).Parent.Name & "<P>"
'Response.write n(0).Clone(1).Text & "<P>"
'Set n = Nothing
'Ebody.Xml("name")(0).RemoveAttr("alias")
'Response.write Ebody.HtmlEncode(Ebody.Xml("name")(0).Xml)
'Ebody.Xml("site")(1).Clear
'Response.write Ebody.HtmlEncode(Ebody.Xml("site")(1).Xml)


'Response.write "--------" & "<P>"
'Response.write "属性控制（有错误）：" & "<P>"
'Response.write Ebody.HtmlEncode(TypeName(Ebody.Xml("site")(0).Parent.Parent.Dom))
'Ebody.Xml("url").Remove
'Ebody.Xml("name").Attr("alias") = Null
'Ebody.Xml("microblog").Remove
'Response.write Ebody.Xml.Sel("//site").Length & "<P>"
'Response.write Ebody.Xml.Select("//site").Length & "<P>"
'Response.write Ebody.Xml("site").Length & "<P>"
'Response.write Ebody.Xml("site").Type & "<P>"
'Ebody.Xml("url")(2).Value = "http://sss.com"
'Response.write TypeName(n) & "<P>"

Response.write "--------" & "<P>"
Response.write "替换节点（有错误）：" & "<P>"
'替换节点
'Set n = Ebody.Xml("name")(1).ReplaceWith(Ebody.Xml.Create("abbr cdata","This is a <b>word</b>."))
'Set n = Ebody.Xml("name").ReplaceWith(Ebody.Xml.Create("abbr cdata","This is a <b>word</b>."))
'Set n = Ebody.Xml("name")(1).ReplaceWith(Ebody.Xml("url")(2))
'Response.write Ebody.HtmlEncode(n.Xml)
'清空
'Ebody.Xml("url").Empty
'Ebody.Xml("name").Clear

Response.write "--------" & "<P>"
Response.write "从前面加入节点：" & "<P>"
'从前面加入节点
Set n = Ebody.Xml("account")(1).Before(Ebody.Xml.Create("abbr cdata","This is a <b>word</b>."))
Set n = Ebody.Xml("account")(1).Before(Ebody.Xml("url")(2))
Set n = Ebody.Xml("account").Before(Ebody.Xml.Create("abbr cdata","This is a <b>word</b>."))
Set n = Ebody.Xml("account").Before(Ebody.Xml("url")(2))

Response.write Ebody.HtmlEncode(n.Xml)
Response.write Ebody.HtmlEncode(Ebody.Xml.Dom.Xml)

'Response.write "--------" & "<P>"
'Response.write "从后面加入节点：" & "<P>"
'
''从后面加入节点
'Set n = Ebody.Xml("account")(2).After(Ebody.Xml.Create("abbr cdata","This is a <b>word</b>."))
'Set n = Ebody.Xml("last")(1).After(Ebody.Xml.Create("abbr cdata","This is a <b>word</b>."))
'Set n = Ebody.Xml("account")(1).After(Ebody.Xml("url")(2))
'Set n = Ebody.Xml("account").After(Ebody.Xml.Create("abbr cdata","This is a <b>word</b>."))
'Set n = Ebody.Xml("account").After(Ebody.Xml("url")(2))

'Response.write Ebody.HtmlEncode(n.Xml)
'Response.write Ebody.HtmlEncode(Ebody.Xml.Dom.Xml)


Response.write "--------" & "<P>"
Response.write "节点取值：" & "<P>"

Response.write Ebody.HtmlEncode(Ebody.Xml("name").Length) & "<P>"
Response.write Ebody.HtmlEncode(Ebody.Xml("site name").Length) & "<P>"
Response.write Ebody.HtmlEncode(Ebody.Xml("site>name").Length) & "<P>"
Response.write Ebody.HtmlEncode(Ebody.Xml("name[alias='Tencent'],url").Length) & "<P>"
Response.write Ebody.HtmlEncode(Ebody.Xml("name[alias='Tencent'],url").Text) & "<P>"
Response.write Ebody.HtmlEncode(Ebody.Xml.Select("//account[@nick='user' and position()<2]").Length) & "<P>"
Response.write Ebody.HtmlEncode(Ebody.Xml.Select("//account[@nick='user' and position()<2]").Xml) & "<P>"
Response.write Ebody.HtmlEncode(Ebody.Xml("account[nick='user'][for!='me'],account[nick!='user']").Xml) & "<P>"

Response.write Ebody.HtmlEncode(Ebody.Xml("site")(1).Find("account").Root.TypeString) & "<P>"
Response.write Ebody.HtmlEncode(Ebody.Xml.Root.TypeString) & "<P>"


' 关闭类
ebody.close "xml"


Response.write "<P>------------------------------------<P>"
Response.write "页面执行时间： " & Ebody.ScriptTime & " 秒"

%>