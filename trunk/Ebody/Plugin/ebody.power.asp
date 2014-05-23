<%
'################################################################################
'## ebody.power.asp
'## -----------------------------------------------------------------------------
'## 功能:	系统权限控制
'## 版本:	1.1 Build 130324
'## 作者:	Tony
'## 日期:	2014/03/22
'## 说明:	系统权限控制
'################################################################################

Class ebody_power
	
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
	' 新建类实例
	' 说明: 当需要创建一个新的实例时可使用此方法
	Public Function [New]()
		Set [New] = New ebody_power		
	End Function

	' 取用户ID
	Public Function GetUserId(pUserName)
		Dim lvSQL : lvSQL = "select user_id from sys_user where lcase(user_name) = '" & LCase(pUserName) & "'"
		Ebody.DB.ExecuteSQL(lvSQL)
		GetUserId = Ebody.DB.GetValue(1,1)
	End Function

	' 取角色ID
	Public Function GetRoleId(pRoleName)
		Dim lvSQL : lvSQL = "select role_id from sys_role where lcase(role_name) = '" & LCase(pRoleName) & "'"
		Ebody.DB.ExecuteSQL(lvSQL)
		GetRoleId = Ebody.DB.GetValue(1,1)
	End Function

	' 验证用户是否存在于角色中
	Public Function IsUserInRole(pUserName, pRoleName)
		Dim lvSQL : lvSQL = "select count(1) from sys_role_user [sru],sys_role [sr],sys_user [su] where [sru].valid = true and [sru].role_id = [sr].role_id and [sru].user_id = [su].user_id and [su].valid = true and lcase([su].user_name) = '" & LCase(pUserName) & "' and [sr].valid = true and lcase([sr].role_name) = '" & LCase(pRoleName) & "'"
		Ebody.DB.ExecuteSQL(lvSQL)
		IsUserInRole = Ebody.IIF(Ebody.DB.GetValue(1,1) = 1, TRUE, FALSE)
	End Function

	' 取对像ID
	Public Function GetObjectId(pObjectName)
		Dim lvSQL : lvSQL = "select object_id from sys_object where lcase(object_name) = '" & LCase(pObjectName) & "'"
		Ebody.DB.ExecuteSQL(lvSQL)
		GetObjectId = Ebody.DB.GetValue(1,1)
	End Function

	' 取资源ID
	Public Function GetResourceId(pResourceName)
		Dim lvSQL : lvSQL = "select resource_id from sys_resource where lcase(resource_name) = '" & LCase(pResourceName) & "'"
		Ebody.DB.ExecuteSQL(lvSQL)
		GetResourceId = Ebody.DB.GetValue(1,1)
	End Function

	' 验证对像是否存在于资源中
	Public Function IsObjectInResource(pObjectName, pResourceName)
		Dim lvSQL : lvSQL = "select count(1) from sys_resource_object [sro],sys_resource [sr],sys_object [so] where [sro].valid = true and [sro].resource_id = [sr].resource_id and [so].valid = true and [sro].object_id = [so].object_id and lcase([so].object_name) = '" & LCase(pObjectName) & "' and [sr].valid = true and lcase([sr].resource_name) = '" & LCase(pResourceName) & "'"
		Ebody.DB.ExecuteSQL(lvSQL)
		IsObjectInResource = Ebody.IIF(Ebody.DB.GetValue(1,1) = 1, TRUE, FALSE)
	End Function



	' 取得用户权限
	' 权限管控重要方法
	' 说明: 有权必须对像生效,对象在指定用户组中,并且有相应的权限
	'       采用乐观权限控制,当object无权时,而对应的resource有权,则object同样有权
	' 返回: true 有权, false 无权
	Public Function CheckPower(pUserName,pObjectName,pActionCode)
		Dim lvSQL, lvPowerValue
		' combine_type 1=资源, 2=对像		

		' 先取object权限
		lvSQL = "select count(1) " &_
				"from " &_
				"sys_role_user [sru], " &_
				"sys_user [su] " &_
				"where 1 = 1 " &_
				"and exists " &_
				"( " &_
				"select 1 from " &_
				"sys_action [sa], " &_
				"sys_power [sp], " &_
				"sys_object [so] " &_
				"where 1 = 1 " &_
				"and [sa].valid = true " &_
				"and [sa].action_id = [sp].action_id " &_
				"and [sa].action_code = '" & pActionCode & "' " &_
				"and [sp].valid = true " &_
				"and [sp].power_flag = true " &_
				"and [sp].role_id = [sru].role_id " &_			
				"and [sp].combine_id = [so].object_id " &_
				"and [sp].combine_type = 2 " &_
				"and [so].valid = true " &_
				"and [so].object_name = '" & pObjectName & "' " &_
				") " &_
				"and [sru].valid = true " &_
				"and [sru].user_id = [su].user_id " &_
				"and [su].user_name = '" & pUserName & "' "

		' 执行SQL,得到权限值
		Ebody.DB.ExecuteSQL(lvSQL)
		lvPowerValue = Ebody.DB.GetValue(1,1)

		' 如果object无权限,则取resource权限
		' 当object挂在resource下,而resource有权,那么此对像同样有权(此为乐观权限控制)
		If lvPowerValue = 0 Then
		
			lvSQL = "select count(1) " &_
					"from " &_
					"sys_role_user [sru], " &_
					"sys_user [su] " &_
					"where 1 = 1 " &_
					"and exists " &_
					"( " &_
					"select 1 from " &_
					"sys_action [sa], " &_
					"sys_power [sp], " &_
					"sys_resource_object [sro], " &_
					"sys_object [so] " &_
					"where 1 = 1 " &_
					"and [sa].valid = true " &_
					"and [sa].action_id = [sp].action_id " &_
					"and [sa].action_code = '" & pActionCode & "' " &_
					"and [sp].valid = true " &_
					"and [sp].power_flag = true " &_
					"and [sp].role_id = [sru].role_id " &_
					"and [sp].combine_id = [sro].resource_id " &_
					"and [sp].combine_type = 1 " &_
					"and [sro].valid = true " &_
					"and [sro].object_id = [so].object_id " &_
					"and [so].valid = true " &_
					"and [so].object_name = '" & pObjectName & "' " &_
					") " &_
					"and [sru].valid = true " &_
					"and [sru].user_id = [su].user_id " &_
					"and [su].valid = true " &_
					"and [su].user_name = '" & pUserName & "' "

			' 执行SQL,得到权限值
			Ebody.DB.ExecuteSQL(lvSQL)
			lvPowerValue = Ebody.DB.GetValue(1,1)
		End If

		' 综合对像,资源的权限得出最终权限
		CheckPower = Ebody.IIF(lvPowerValue = 1, TRUE, FALSE)
	End Function

	' 用户登陆
	Public Function Login(pUserName,pUserPass)
		Dim lvLogingValue
		Ebody.DB.ExecuteSQL("select count(1) from sys_user where lcase(user_name) = '" & LCase(pUserName) & "' and password = '" & pUserPass & "'")
		Login = Ebody.IIF(Ebody.DB.GetValue(1,1) = 1, TRUE, FALSE)
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