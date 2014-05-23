<%
'######################################################################
'## ebody.config.asp
'## -------------------------------------------------------------------
'## Ebdoy 配置文件
'######################################################################

' 必须正确设置'Ebody.asp'文件在网站中的路径，以"/"开头，路径基于网站根目录
Ebody.BasePath = "/Ebody"

' 设置文件编码(通常为'GBK'或者'UTF-8')
Ebody.CharSet = "UTF-8"

' 设置如何处理载入的UTF-8文件的BOM信息(keep保留/remove移除/add新增)
Ebody.FileBOM = "remove"

' 是否加密Cookies数据(true是/false否)
Ebody.CookieEncode = False

' 设置FSO组件的名称(如果服务器上修改过)
Ebody.FsoName = "Scripting.FileSystemObject"

' 配置默认数据库连接字符串
' Access:
Ebody.DBConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ebody.MapPath("/Ebody/db/ebody.mdb") & ";Jet OLEDB:Database Password=;"

' Oracle:
'Ebody.DBConnStr = "Provider=MSDAORA;Data Source=(DESCRIPTION = (ADDRESS_LIST = (ADDRESS = (PROTOCOL = TCP)(Host = 192.168.0.34)(Port = 1522))) (CONNECT_DATA =(SID = ERPDEV))); User Id=apps; Password=apps"
%>