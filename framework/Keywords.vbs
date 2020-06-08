class[action name]

	Public dt
	Public Iteration
	Public Status
	Public Stepnum
	Public object_id
	Public value
	Public result
	Public store_result_as

	Private Sub Class_Initialize()
		Call InfoClassInstance(me, C_OBJ_LOADED_MSG)
	End Sub
	
	Private Sub Class_Terminate()
		Call InfoClassInstance(me, C_OBJ_UNLOADED_MSG)
	End Sub
		
	Sub Run()
		' your code here processing object_id and  value into what you need
		
        On Error Goto 0
		Status = 0
	End Sub
	
	' Other fucntions and subs to ease maintenance
	
end class