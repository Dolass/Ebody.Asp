<% @ language="vbscript" %>

<% 
' ֱ�ӵ������������Ȩ�޿��ƣ�һ��ʹ��Regsvr32.exe E:\Tony\Ebody101\demo.dll��ע�������ʹ��Regsvr32.exe /u E:\Tony\Ebody101\demo.dll��ж�����
%>

<% ' ����Ϊֱ������DLLԴ�ļ������ɲ���ע�� %>
<!--METADATA TYPE="typelib" FILE="H:\WEB\Taoya2014\Core\Ebody111\Demo\DLL\demo.dll" -->

<%
Dim oDemoDll
Set oDemoDll=Server.Createobject("demo.show") 
oDemoDll.Msg("<font size=4>this is my first test dll</font><input type=text value=test>")
Set oDemoDll=Nothing 
%>