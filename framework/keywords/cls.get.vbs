Class [get]

	Public dt
	Public Iteration
	Public Status
	Public Stepnum
	Public object_id ' schema to split on | or ,?
	Public value     ' schema to split on | or ,?
	Public result

	Sub Class_Initialize()
		InfoClassInstance(me, "Loaded successfully")
	End sub
	
	Private Sub Class_Terminate()
		InfoClassInstance(me, "Terminated successfully")
	End Sub

	function run( object )
		print("get value")
		run = "value)
	End Function
	
End Class
