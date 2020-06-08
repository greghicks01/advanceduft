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
		InfoClassInstance(me, "Loaded successfully")
	End Sub
	
	Private Sub Class_Terminate()
	End Sub
		
	Sub Run()
		' your code here processing object_id and  value into what you need
		
        On Error Goto 0
		Status = 0
	End Sub
	
	' Other fucntions and subs to ease maintenance
	
end class