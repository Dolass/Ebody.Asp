<%
'################################################################################
'## ebody.music.asp
'## -----------------------------------------------------------------------------
'## 功能:	类的内容主题
'## 版本:	1.1 Build 130324
'## 作者:	Tony
'## 日期:	2013/03/24
'## 说明:	类的主要功能描述
'################################################################################

Class ebody_music
	
'================================================================================
'== Variable
'================================================================================


'--------------------------------------------------------------------------------
'-- Public
'--------------------------------------------------------------------------------


'--------------------------------------------------------------------------------
'-- Private
'--------------------------------------------------------------------------------


'================================================================================
'== Event
'================================================================================

	Private Sub Class_Initialize()
		'On Error Resume Next
	End Sub

	Private Sub Class_Terminate()
		'On Error Resume Next
	End Sub

'================================================================================
'== Public
'================================================================================


'--------------------------------------------------------------------------------
'-- Let Property
'--------------------------------------------------------------------------------


'--------------------------------------------------------------------------------
'-- Get Property
'--------------------------------------------------------------------------------


'--------------------------------------------------------------------------------
'-- Function
'--------------------------------------------------------------------------------
Public Function show()
	Response.write "当前执行的是: ebody.music.asp类，使用的是Extend扩展类方法。"
End Function

'--------------------------------------------------------------------------------
'-- Sub
'--------------------------------------------------------------------------------


'================================================================================
'== Private
'================================================================================


'--------------------------------------------------------------------------------
'-- Function
'--------------------------------------------------------------------------------


'--------------------------------------------------------------------------------
'-- Sub
'--------------------------------------------------------------------------------


End Class
%>