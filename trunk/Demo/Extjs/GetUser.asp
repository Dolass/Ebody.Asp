<!--#include file="../../Ebody/ebody.asp"-->

<%
' ===============================================
' 创建基类
' ===============================================
'Dim Ebody : Set Ebody = New Ebody_Base
Ebody.CharSet = "UTF-8"

' ===============================================
' 演示 3
' 目的: DB json快捷输出
' ===============================================

' 载入类
Ebody.Use "db"

' 设定DB连接字串
Ebody.db.ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ebody.MapPath("../../db/ebody.mdb") & ";Jet OLEDB:Database Password=;"

' json生成代码
Ebody.db.Open
Ebody.db.ExecuteSQL "select user_id, user_name from sys_user order by user_id desc"

' 以下为不显示总条数的json
Response.write "{" & Ebody.UnEscape(Ebody.db.json("items", Ebody.db.rs)) & "}"

' 关闭类
Ebody.Close "db"

' ===============================================
' 销毁基类
' ===============================================
Set Ebody = Nothing

%>