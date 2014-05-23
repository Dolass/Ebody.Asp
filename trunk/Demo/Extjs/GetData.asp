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

' 取数据总数
Ebody.db.ExecuteSQL("select count(*) from sys_user")
Dim lvRecordCount : lvRecordCount = Ebody.db.getvalue(1,1)

' 以下为正常取数，不含分页
'Ebody.db.ExecuteSQL("select user_id, user_name, password, desc, valid from sys_user")

' 页记录数(extjs会传入start,limit两个参数,start=开始记录行,limit=每页显示记录数,相当于pagesize)
' 取得SQL,依据extjs的pagesize来自动设定后台的读取量,做到按需读取记录
Dim lvLimit, lvStart, lvSQL
lvLimit = Request("limit")
lvStart = Request("start")
If lvLimit = "" Then lvLimit = 5
If lvStart = "" Then lvStart = 0

' 以下加入access数据的分页功能
' 注意: 栏位一定要小写,否则extjs认不到
' 依rownum来限定读取量
Ebody.db.ExecuteSQL "select user_id, user_name, password, desc, sex, born_date, valid from (select * , (select count(*)+1 from (select * from sys_user) a where a.user_id <b.user_id) as rownum from (select * from sys_user) b) c where c.rownum > " & lvStart & " and c.rownum <= " & cstr(cint(lvStart) + cint(lvLimit)) & " order by user_id desc"

' 以下为显示总条数的json
'Response.write "{total:" & lvRecordCount & "," & Ebody.UnEscape(Ebody.db.json("items", Ebody.db.rs)) & "}"
'Response.write "{total:" & ebody.db.rs.RecordCount & "," & Ebody.UnEscape(Ebody.db.json("items", Ebody.db.rs)) & "}"

' 以下为不显示总条数的json
'Response.write "{" & Ebody.UnEscape(Ebody.db.json("items", Ebody.db.rs)) & "}"

' 以下为固定输出，以做对比验证
'Response.write "{total:7,items:[{user_id:1,user_name:""admin"",password:""admin"",desc:""管理员"",valid:""True""},{user_id:2,user_name:""demo"",password:""demo"",desc:""演示帐户"",valid:""False""}]}"




' -------------------------------------------
' 使用json类来生成数据
' 更为灵活，可用于改变内部数据
ebody.use "json"
ebody.use "aes"

Set rs = Ebody.db.rs.Clone
Set o = Ebody.Json.New(0)

jName = "items"
total = 0

' 依据rs结构，存入json对像
If Ebody.Has(rs) Then
	total = rs.RecordCount
	' 定义total
	o("total") = lvRecordCount

	' 定义items内容
	o(jName) = Ebody.Json.New(1)
	While Not rs.Eof
		o(jName)(Null) = Ebody.Json.New(0)
		For Each fi In rs.Fields
			o(jName)(Null)(fi.Name) = ebody.iif(fi.name="password",ebody.aes.encode(fi.Value),fi.value)	'针对密码加密显示
		Next
		rs.MoveNext
	Wend
End If

ebody.close "json"
ebody.close "aes"

' 以下为不显示总条数的json
Response.write Ebody.UnEscape(o.JSON)
'----------------------------------------------


Ebody.db.Close

' 关闭类
Ebody.Close "db"

' ===============================================
' 销毁基类
' ===============================================
Set Ebody = Nothing

%>