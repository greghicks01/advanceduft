Class [Close Browser]

	Public dt
	Public Iteration
	Public Status
	Public Stepnum
	Public Details
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
		me.Status.[+=] CloseBrowser()
		'-----
        Call ReportActionStatus(me)
	End Sub
	
	Function CloseBrowser()
	
		CloseBrowser = 0
		
		PrintReportInfo "Step " & me.StepNum, " INFO: Closing browser"
		
		On error resume next
			systemutil.CloseProcessByName("Chrome.exe")
			systemutil.CloseProcessByName("Internet Explorer.exe")
			systemutil.CloseProcessByName("Edge.exe")
			systemutil.CloseProcessByName("firefox.exe")
			systemutil.CloseProcessByName("")
			systemutil.CloseProcessByName("")
			systemutil.CloseProcessByName("")
		On error goto 0
			
		If err.number <> 0  Then ' catch
		
				CloseBrowser = 1
				
		End If
			
	End Function
		
End Class
