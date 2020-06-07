Class [Login]

	Public dt
	Public Iteration
	Public Status
	Public Stepnum
	Public object_id ' schema to split on | or ,?
	Public value     ' schema to split on | or ,?
	Public result
	
	Sub run()
	
		Dim oAction
		
		ASSERT_RESULT(GetClassInstance(oAction, "[" & "set" & "]"))
		oAction.object_id = ""
		oAction.value = ""
		oAction.run()
		oAction.run()
		oAction = Nothing
		
	
		ASSERT_RESULT(GetClassInstance(oAction, "[" & "click" & "]"))
		oAction.run()
		oAction = Nothing
		
	End Sub
	
End Class 

	