class[math]

	Public dt
	Public Iteration
	Public Status
	Public Stepnum
	Public object_id
	Public value
	Public result

	Sub Class_Initialize()
		Call InfoClassInstance(me, C_OBJ_LOADED_MSG)
	End sub
	
	Private Sub Class_Terminate()
		Call InfoClassInstance(me, C_OBJ_UNLOADED_MSG)
	End Sub
	
	Sub Run()
		' your code here processing value into what you need
		' takes value and tries to perform an eval
		Status = 0
		on Error Resume Next
		result = Eval(value)
		If Err.Number > 0 then Status = 1: Exit Sub
	End Sub
	
	' Other fucntions and subs to ease maintenance
	
end Class

'value = "45+10"
'result = Eval(value)
'Debug.WriteLine(result)
'System.Collections.ArrayList
