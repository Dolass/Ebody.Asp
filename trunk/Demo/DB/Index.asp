<!--#include file="../../Ebody/ebody.asp"-->

<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" type="text/css" href="Style/EasyData/Base.css" />
<title>DB</title>
</head>

<%
' ===============================================
' 创建基类
' ===============================================
'Dim Ebody : Set Ebody = New Ebody_Base
Ebody.CharSet = "UTF-8"

' ===============================================
' 演示 1
' 目的: 主要演示基本功能及属性
' ===============================================

Response.write "<p>"
Response.write "###############################" & "<br>"
Response.write "## 主要演示基本功能及属性" & "<br>"
Response.write "###############################" & "<p>"

' 载入类
Ebody.Use "db"

Ebody.db.ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ebody.MapPath("../../db/ebody.mdb") & ";Jet OLEDB:Database Password=;"
Ebody.db.PageSize = 5
Ebody.db.PageIndex = 1
Ebody.db.Open
Ebody.db.ExecuteSQL("select user_id, user_name from sys_user")

Response.write "---类属性---" & "<p>"
Response.write "SQL,取得执行过的最后一条SQL语句: " & ebody.db.sql & "<p>"
Response.write "CONNStr,取得数据连接字符串: " & ebody.db.ConnStr & "<p>"
Response.write "PageCount,取分页总数: " & ebody.db.PageCount & "<p>"

Response.write "---基础功能类函数---" & "<p>"
'Response.write "Json,转换成Json格式代码: " & Ebody.db.Json & "<P>"

Response.write "---验证类函数(Is打头)---" & "<p>"
Response.write "IsOpen,判断对像是否已经打开: " & ebody.db.IsOpen(Ebody.db.RS) & "<p>"

Response.write "---系统取值类函数(Get打头)---" & "<p>"
Response.write "GetValue,取得指定行列坐标的值: " & Ebody.db.GetValue(3,2) & "<P>"
Response.write "GetColName,取得指定列名称: " & Ebody.db.GetColName(2) & "<P>"

Ebody.db.Close

' 关闭类
Ebody.Close "db"

' ===============================================
' 演示 2
' 目的: 主要演示数据库连接取值的使用方法
' ===============================================

Response.write "<p>"
Response.write "###############################" & "<br>"
Response.write "## 主要演示数据库连接取值的使用方法" & "<br>"
Response.write "###############################" & "<p>"
Response.write "<p>"

' 载入类
Ebody.Use "db"

' 设定DB连接字串
Ebody.db.ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ebody.MapPath("../../db/ebody.mdb") & ";Jet OLEDB:Database Password=;"
Ebody.db.Open
Ebody.db.ExecuteSQL("select user_id, user_name from sys_user")
Response.write "GetValue: " & Ebody.db.GetValue(3,2) & "<P>"
Response.write "GetColName: " & Ebody.db.GetColName(2) & "<P>"
Ebody.db.Close

' 关闭类
Ebody.Close "db"

' ===============================================
' 演示 2
' 目的: 主要演示多数据连接取值的方法
' ===============================================

Response.write "<p>"
Response.write "###############################" & "<br>"
Response.write "## 主要演示多数据连接取值的方法" & "<br>"
Response.write "###############################" & "<p>"
Response.write "<p>"

' 载入类
Ebody.Use "db"

' 通过DB的NEW方法来创建一个新的DB对像
Set loDB = ebody.db.New
'loDB.ConnStr = "Provider=MSDAORA;Data Source=(DESCRIPTION = (ADDRESS_LIST = (ADDRESS = (PROTOCOL = TCP)(Host = 192.168.0.34)(Port = 1522))) (CONNECT_DATA =(SID = ERPDEV))); User Id=apps; Password=apps"
loDB.SetConn "oracle", "erpdev", "", "192.168.0.34", "1522", "apps", "apps"

loDB.PageSize = 3
loDB.PageIndex = 2

loDB.Open
loDB.ExecuteSQL("select user_id, user_name from fnd_user")
Response.write loDB.GetValue(4,2)
loDB.Close
Set loDB = Nothing

' 关闭类
Ebody.Close "db"


' ===============================================
' 演示 3
' 目的: json输出
' ===============================================

Response.write "<p>"
Response.write "###############################" & "<br>"
Response.write "## json输出" & "<br>"
Response.write "###############################" & "<p>"
Response.write "<p>"

' 载入类
Ebody.Use "db"

' 设定DB连接字串
Ebody.db.ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ebody.MapPath("../../db/ebody.mdb") & ";Jet OLEDB:Database Password=;"

' 设定json生成参数
'Ebody.db.json.QuotedVars = false

' json生成代码
Ebody.db.Open
Ebody.db.ExecuteSQL("select top 5 resource_id, resource_name from sys_resource")
Response.write "Json输出: " & Ebody.UnEscape(Ebody.db.json("resource", Ebody.db.rs))
Ebody.db.Close

' 关闭类
Ebody.Close "db"

' ===============================================
' 销毁基类
' ===============================================
Set Ebody = Nothing

%>