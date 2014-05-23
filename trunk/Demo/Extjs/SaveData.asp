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

' 更新数据
Dim lvId, lvField, lvValue

lvId = Request("id")
lvField = Request("name")
lvValue = Request("value")

'Ebody.db.ExecuteSQL("update sys_user set " & lvField & " = " & lvValue & " where user_id = " & lvId)
Ebody.db.ExecuteSQL("update sys_user set user_name " =  & Request("user_name") & " where user_id = " & Request("user_id"))

' 关闭类
Ebody.Close "db"

' ===============================================
' 销毁基类
' ===============================================
Set Ebody = Nothing

%>