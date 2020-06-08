Class StopRunSession
    ' --------------------------------------------------
    ' Reusable Action: StopRunSession
    ' Description: Stops the run in case of an unhandled Error/exception
    ' --------------------------------------------------
    Public Status
    Public Iteration
    Public StepNum
    Public dt
    Public Details
    
    Public Function Run()
        me.Details = "Ended with "
        me.Status.[=]Reporter.RunStatus
        '--- Report
        Call ReportActionStatus(me)
        '--- Stops the run session
        ExitTest(Reporter.RunStatus)
    End Function
    
    Private Sub Class_Initialize
        Call InfoClassInstance(me, C_OBJ_LOADED_MSG)
        Set me.Status = [As Num](0)
    End Sub
    
    Private Sub Class_Terminate
        Call InfoClassInstance(me, C_OBJ_UNLOADED_MSG)
        Set me.Status = Nothing
    End Sub
End Class

Class EventHandler
    Function RunMappedProcedure(ByVal strError)
        Dim oProcedure
        '--- Try to execute the procedure associated with the error (if exists)
        If GetClassInstance(oProcedure, Environment("ERR_" & CStr(Abs(strError)))) = 0 Then
            RunMappedProcedure = oProcedure.Run
            Exit Function
        End If
        '--- Try to execute the default procedure to handle errors (if exists)
        If GetClassInstance(oProcedure, Environment("DEFAULT_ERROR_HANDLER")) Then
            RunMappedProcedure = oProcedure.Run
            Exit Function
        End If
    End Function
End Class

Class ClearError
    ' --------------------------------------------------
    ' Reusable Action: ClearError
    ' Description: Clears the error in case of an unhandled Error/exception
    ' --------------------------------------------------
    Public Status
    Public Iteration
    Public StepNum
    Public dt
    Public Details
    Public Function Run()
        me.Details = "Ended with "
        me.Status.[=]0
        '--- Report
        Call ReportActionStatus(me)
        '--- Clears the error
        Err.Clear
    End Function
    Private Sub Class_Initialize
        Call InfoClassInstance(me, C_OBJ_LOADED_MSG)
        Set me.Status = [As Num](0)
    End Sub
    Private Sub Class_Terminate
        Call InfoClassInstance(me, C_OBJ_UNLOADED_MSG)
        Set me.Status = Nothing
    End Sub
End Class