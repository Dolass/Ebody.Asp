<!--#include file="../../Ebody/ebody.asp"-->

<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" type="text/css" href="Style/EasyData/Base.css" />
<title>Json</title>
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

' 载入类
Ebody.Use "json"

Response.write "---类属性设定---" & "<p>"
'Ebody.json.QuotedVars = True
'Ebody.json.StrEncode = True

Response.write "---类属性---" & "<p>"
'Response.write "HttpType,取得服务器验证类型: " & Ebody.HttpType & "<p>"

Response.write "---基础功能类函数---" & "<p>"
'Response.write "IIF,判断三元表达式: " & Ebody.IIF(1=2, "YES", "NO") & "<p>"

Response.write "---验证类函数(Is打头)---" & "<p>"
'Response.write "IsExists,检测文件或文件夹或驱动器(磁盘)是否存在: "		& Ebody.IsExists("e:\ebody\demo.asp") & "<p>"

Response.write "---系统取值类函数(Get打头)---" & "<p>"
'Response.write "GetPathAbs,依相对路径取得文件绝对物理路径: "		& Ebody.GetPathAbs("/core/class") & "<p>"

' 关闭类
Ebody.Close "json"

' ===============================================
' 演示 2
' 目的: 功能的应用演示
' ===============================================

Response.write "<p>"
Response.write "###############################" & "<br>"
Response.write "## 自定义json输出演示" & "<br>"
Response.write "###############################" & "<p>"
Response.write "<p>"

Ebody.Use "json"
Ebody.Use "db"

'Ebody.db.ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ebody.MapPath("../../db/ebody.mdb") & ";Jet OLEDB:Database Password=;"
Ebody.db.Open
Ebody.db.ExecuteSQL("select top 2 user_id, user_name from sys_user")

Set rs = Ebody.db.rs.Clone
Set o = Ebody.Json.New(0)

jName = "abc"

		'If notjs Then o.StrEncode = False
		total = 0

		If Ebody.Has(rs) Then
			total = rs.RecordCount

Response.write "total:" & total & "<P>"

			'If Ebody.Has(totalName) Then o(totalName) = total
			o(jName) = Ebody.Json.New(1)

Response.write "total:" & total & "<P>"

			While Not rs.Eof

				o(jName)(Null) = Ebody.Json.New(0)

				For Each fi In rs.Fields

Response.write fi.name & "=" & fi.Value & "<P>"

					o(jName)(Null)(fi.Name) = fi.Value
				Next
				rs.MoveNext
			Wend
		End If

		'tmpStr = o.JSON


'Dim o : Set o = Ebody.Json.New(0)
'o("abc") = Ebody.Json.New(1)
'o("abc")(Null) = Ebody.Json.New(0)

'Dim o : Set o = Ebody.Json.New(0)
Response.write "Json输出: " & o.JSON
Response.write "<p>"

'Response.write Ebody.json.ToJSON("abc:123:测试") & "<p>"

Ebody.Close "json"
Ebody.Close "db"
Response.write "<p>"


' ===============================================
' 演示 3
' 目的: DB json快捷输出
' ===============================================

Response.write "<p>"
Response.write "###############################" & "<br>"
Response.write "## DB json快捷输出" & "<br>"
Response.write "###############################" & "<p>"
Response.write "<p>"

' 载入类
Ebody.Use "db"

' 设定DB连接字串
'Ebody.db.ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ebody.MapPath("../../db/ebody.mdb") & ";Jet OLEDB:Database Password=;"

' 设定json生成参数
'Ebody.db.json.QuotedVars = false

' json生成代码
Ebody.db.Open
Ebody.db.ExecuteSQL("select top 2 user_id, user_name from sys_user")
Response.write "Json输出: " & Ebody.UnEscape(Ebody.db.json("abc", Ebody.db.rs))
Ebody.db.Close

' 关闭类
Ebody.Close "db"
Response.write "<p>"

' ===============================================
' 演示 4
' 目的: 后台解析json数据
' ===============================================

Response.write "<p>"
Response.write "###############################" & "<br>"
Response.write "## 后台解析json数据" & "<br>"
Response.write "###############################" & "<p>"
Response.write "<p>"


Ebody.Use "json"

Dim JsonStr, objTest
JsonStr = "{name:""alonely"", age:24, email:[""ycplxl1314@163.com"",""ycplxl1314@gmail.com""], family:{parents:[""父亲"",""母亲""],toString:function(){return ""家庭成员"";}}}"
Set objTest = Ebody.json.parse(JsonStr)

' 取json数据
Response.write objTest.name & "的邮件地址是：" & Ebody.json.getValue(objTest.email,1) & ",共有邮件地址" & objTest.email.length & "个" & "<P>"
Response.write objTest.name & "家庭成员是：" & Ebody.json.getValue(objTest.family.parents,0) & "<P>"

' 遍历json数据
Response.write "家庭成员有：" & "<P>"
For Each value In objTest.family.parents
	Response.write value & "<P>"
Next

Ebody.Close "json"

' ===============================================
' 演示 5
' 目的: 后台解析json数据
' ===============================================

Response.write "<p>"
Response.write "###############################" & "<br>"
Response.write "## 后台解析json数据" & "<br>"
Response.write "###############################" & "<p>"
Response.write "<p>"

Ebody.Use "json"
Ebody.Use "db"

' json生成代码
' 设定DB连接字串
'Ebody.db.ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ebody.MapPath("../../db/ebody.mdb") & ";Jet OLEDB:Database Password=;"
Ebody.db.SetConn "oracle", "erpdev", "", "192.168.0.34", "1522", "apps", "apps"
Ebody.db.Open
'Ebody.db.ExecuteSQL("select user_id, user_name from fnd_user where rownum < 10")
'Ebody.db.ExecuteSQL("select top 2 user_id, user_name from sys_user")
ebody.db.executeSql("select a.party_name as user_name, a.email_address as email from hz_parties a where rownum < 10 and a.email_address is not null")
JsonStr = "{action:""query""," & Ebody.UnEscape(Ebody.db.json("abc", Ebody.db.rs)) & "}"
Ebody.db.Close

' 分析json字串，返回为json对像
Set objTest = Ebody.json.parse(JsonStr)

' 取单个数据
'{"abc":[{"user_id":1,"user_name":"admin"},{"user_id":2,"user_name":"demo"}]}
Response.write "数据条数：" & objTest.abc.length & "<P>"
Response.write "动作：" & objTest.action & "<P>" 

' 遍历json中的数组数据
Response.write "遍历Json数据" & "<P>"
For Each value In objTest.abc
	Response.write "用户名：" & value.user_name & " 邮件：" & value.email & "<P>"
Next

Ebody.Close "json"
Ebody.Close "db"


' ===============================================
' 注销类
' ===============================================
Set Ebody = Nothing

%>