class[math]

	Public dt
	Public Iteration
	Public Status
	Public Stepnum
	Public object_id
	Public value
	Public result
	Public store_result_as

	Sub Class_Initialize()
		Call InfoClassInstance(me, C_OBJ_LOADED_MSG)
		Set me.Status = [as Num](0)
		me.Details = "Ended with "
	End sub
	
	Private Sub Class_Terminate()
		Call InfoClassInstance(me, C_OBJ_UNLOADED_MSG)
		Set me.Status = Nothing
	End Sub
	
	Sub Run()
		' your code here processing value into what you need
		' takes value and tries to perform an eval
		me.Status.[+=]performMath()
		'-----
        Call ReportActionStatus(me)
	End Sub
	
	' Other fucntions and subs to ease maintenance
	Function performMath()
		On Error Resume Next
			me.return = Eval(value)
		If Err.Number <> 0 Then performMath = Err.Number
			
	End Function
	
	
end Class

'value = "45+10"
'result = Eval(value)
'Debug.WriteLine(result)
'System.Collections.ArrayList
