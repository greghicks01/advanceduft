Class [Login]

	Public dt
	Public Iteration
	Public Status
	Public Stepnum
	Public object_id ' schema to split on | or ,?
	Public value     ' schema to split on | or ,?
	Public result
	Public store_result_as

	Sub Class_Initialize()
		Call InfoClassInstance(me, C_OBJ_LOADED_MSG)
		Set me.Status = [As Num](0)
	End sub
	
	Private Sub Class_Terminate()
		Call InfoClassInstance(me, C_OBJ_UNLOADED_MSG)
		Set me.Status = Nothing
	End Sub
	
	Sub run()
	
		Dim oAction
	
		
	End Sub
	
End Class 

	