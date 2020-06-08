Class [get]

	Public dt
	Public Iteration
	Public Status
	Public Stepnum
	Public object_id ' schema to split on | or ,?
	Public value     ' schema to split on | or ,?
	Public result

	Sub Class_Initialize()
		Call InfoClassInstance(me, C_OBJ_LOADED_MSG)
	End sub
	
	Private Sub Class_Terminate()
		Call InfoClassInstance(me, C_OBJ_UNLOADED_MSG)
	End Sub

	sub run( object )
		print("get value")
		result = "value"
		Status = 0
	End Sub
	
End Class
