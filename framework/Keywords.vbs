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
		Set me.Status = [As Num](0)
	End Sub
	
	Private Sub Class_Terminate()
		Call InfoClassInstance(me, C_OBJ_UNLOADED_MSG)
		Set me.status = Nothing
	End Sub
		
	Sub Run()
		me.Details = "Ended with "
		' your code here processing
		me.Status.[+=]EnterUserName()
		'-----
        Call ReportActionStatus(me)
	End Sub
	
	' Other functions and subs to ease maintenance
	
end class