Class autoload

	Private projectroot
		
	Sub Class_Initialize()
	
		If qcutil.IsConnected Then
			projectroot = "[ALM\Resources] Resources\Automation"
		else
			projectroot = Environment.Value("PROJECT_ROOT" )
		End If
		
		Call LocalFileList()
		
	End Sub
	
	Sub Class_Terminate()
		
	End Sub
	
	Private Function LocalFileList()
	
		Set fso = CreateObject("Scripting.FileSystemObject")
		
		For each Folder in fso.GetFolder(projectroot).SubFolders
		
			For each File in Folder.Files
			
				ext = lCase(fso.GetExtensionName(File))
				
				If  (ext = "vbs" or ext = "qfl" )  Then
					include File.path
				End If
				
			Next
		Next
		
	End Function
	
	Function include(filename)
	
		print "loading " + filename
		
		On error resume next
		ExecuteFile filename
		On error goto 0
		
		If err.number <>0 Then
			print "failed to load " + filename
		End If
		
	End Function
		
End Class

set al = new autoload
RunTest()

