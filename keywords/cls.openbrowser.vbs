' See https://admhelp.microfocus.com/uft/en/15.0-15.0.1/UFT_Help/Subsystems/OMRHelp/OMRHelp.htm#Web/Utility_WebUtil.htm#LaunchBrowser
' set up your browser or device in value

Class [Open Browser]
	
	Public dt
	Public Iteration
	Public Status
	Public Stepnum
	Public Details
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
		me.Status.[+=]OpenBrowser()
		'-----
        Call ReportActionStatus(me)
	End Sub
	
	' See https://admhelp.microfocus.com/uft/en/15.0-15.0.1/UFT_Help/Subsystems/OMRHelp/OMRHelp.htm#Web/Utility_WebUtil.htm#LaunchBrowser for all values
	Function OpenBrowser()
	
		set parameters = CreateObject("System.Collections.ArrayList")
		
		OpenBrowser = 0
		
		If instr(me.value, ",") <> 0 Then
			For each param in split(value, ",")
				parameters.add chr(34) & param & chr(34)
			Next
			me.value = join(parameters.toarray, ",")
		else
			me.value = chr(34) & me.value & chr(34)
		End If
		
		' -- try
		On Error Resume Next
		
			PrintReportInfo "Step " & me.StepNum, " INFO: Launch " & me.value & " browser"
			tmp = "webutil.LaunchBrowser " & me.value
			print tmp
			execute tmp
			' -- catch
			If Err.number <> 0 Then
				print err.number
				OpenBrowser = err.number
			End If
		
		On error goto 0
		
	End Function
	
End Class

