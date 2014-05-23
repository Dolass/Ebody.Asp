<% @ language="vbscript" %>

<% 
' 直接调用组件，会有权限控制，一般使用Regsvr32.exe E:\Tony\Ebody101\demo.dll来注册组件，使用Regsvr32.exe /u E:\Tony\Ebody101\demo.dll来卸掉组件
%>

<% ' 以下为直接引用DLL源文件，即可不用注册 %>
<!--METADATA TYPE="typelib" FILE="H:\WEB\Taoya2014\Core\Ebody111\Demo\DLL\demo.dll" -->

<%
Dim oDemoDll
Set oDemoDll=Server.Createobject("demo.show") 
oDemoDll.Msg("<font size=4>this is my first test dll</font><input type=text value=test>")
Set oDemoDll=Nothing 
%>