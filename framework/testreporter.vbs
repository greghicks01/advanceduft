Function ReportActionStatus(ByRef p)
    ' --------------------------------------------------
    ' Function : ReportActionStatus
    ' Purpose : Reports an event to the UFT reporter with the data of the referenced Action
    ' Args : ByRef p
    ' Returns : N/A
    ' --------------------------------------------------
    Reporter.ReportEvent GetNormalizedStatus(p.Status.Value),
    TypeName(p), p.Details & GetStatusText(p.Status.Value)
End Function

Function GetStatusText(ByVal iStatus)
    ' --------------------------------------------------
    ' Function : GetStatusText
    ' Purpose : Returns the text associated with a status
    ' Args : ByVal iStatus
    ' Returns : "success", "failure"
    ' --------------------------------------------------
    Dim sStatus
    Select Case CInt(iStatus)
        Case 0, 2, 4 'micPass, micDone, micInfo
        sStatus = "success"
        Case Else 'micFail, micWarning
        sStatus = "failure"
    End Select
    GetStatusText = sStatus
End Function

Function GetNormalizedStatus(ByVal iStatus)
    ' --------------------------------------------------
    ' Function : GetNormalizedStatus
    ' Purpose : Returns the status as 0 or 1
    ' Args : ByVal iStatus
    ' Returns : 0 or 1
    ' --------------------------------------------------
    GetNormalizedStatus = micPass
    If CLng(iStatus) <> CLng(micPass) Then
        GetNormalizedStatus = micFail
    End If
End Function

Function PrintReportInfo(ByVal sSender, ByVal sMessage)
    ' --------------------------------------------------
    ' Function : PrintReportInfo
    ' Purpose : Reports an info event to the UFT reporter and log
    ' Args : ByVal sSender
    ' ByVal sMessage
    ' Returns : N/A
    ' --------------------------------------------------
    Print sSender & ": " & sMessage
    Reporter.ReportEvent micInfo, sSender, sMessage
End Function