<!--#include file="../../Ebody/ebody.asp"-->

<%
'Dim Ebody : Set Ebody = New Ebody_Base
Dim lvUser : lvUser = "guest"
Dim lvObject : lvObject = "personal.asp"
Dim lvAction : lvAction = "new"

Ebody.extend "power"
Ebody.use "db"

Ebody.db.ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ebody.MapPath("../../db/ebody.mdb") & ";Jet OLEDB:Database Password=;"
Ebody.db.open

Response.write "验证信息<P>"
Response.write "用户:" & lvUser & "<P>"
Response.write "对像:" & lvObject & "<P>"
Response.write "操作:" & lvAction & "<P>"
Response.write "验证结果，是否有权:" & ebody.ext.power.CheckPower(lvUser,lvObject,lvAction) & "<P>"
Response.write "--------------------------------------------" & "<P>"
Response.write "取用户ID:" & ebody.ext.power.GetUserId("TONY") & "<P>"
Response.write "取角色ID:" & ebody.ext.power.GetRoleId("GUEST") & "<P>"
Response.write "取对像ID:" & ebody.ext.power.GetObjectId("personal.asp") & "<P>"
Response.write "取资源ID:" & ebody.ext.power.GetResourceId("功能管理") & "<P>"
Response.write "用户登陆:" & ebody.ext.power.Login("tony","ty") & "<P>"
Response.write "--------------------------------------------" & "<P>"
Response.write "验证用户是否存在于角色中:" & ebody.ext.power.IsUserInRole("tony","demo") & "<P>"
Response.write "验证对像是否存在于资源中:" & ebody.ext.power.IsObjectInResource("index.asp","通用功能") & "<P>"

Ebody.close "db"
ebody.ext.close "power"
Set Ebody = Nothing
%>