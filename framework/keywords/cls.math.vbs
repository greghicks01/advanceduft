class[math]

	Public dt
	Public Iteration
	Public Status
	Public Stepnum
	Public object_id
	Public value
	Public result
	
	Sub Run()
		' your code here processing value into what you need
		' takes value and tries to perform an eval
		on Error Resume Next
		result = Eval(value)
		If Err.Number > 0 then Status = 1: Exit Sub
		Status = 0
	End Sub
	
	' Other fucntions and subs to ease maintenance
	
end Class

value = "45+10"
result = Eval(value)
Debug.WriteLine(result)
System.Collections.ArrayList
