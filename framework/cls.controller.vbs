Const C_STR_TEST_DATA_XLS = "testsample.xlsx"
Const C_OBJ_OF_CLASS_MSG = "--- Object of Class "
Const C_OBJ_LOADED_MSG = " was loaded ---"
Const C_OBJ_UNLOADED_MSG = " was unloaded ---"

Class Controller
    
    Public Status
    Public Details
    
    Function Run(ByVal strTestSetsPathName)
        ' -------------------------------------------------
        ' Function : Run
        ' Purpose : Runs the steps (procedures implemented as Command Wrappers)
        ' Args : ByVal strTestSetsPathName
        ' Returns : 0 on success; 1 on failure
        ' -------------------------------------------------
        ' Usage : Run("C:\Automation\Test_Sets\")
        ' Notes : 
        ' 1) Uses a Local DataSheet to control the steps flow
        ' 2) Uses GetClassInstance
        ' 3) Uses CNum
        ' 4) Uses ASSERT_RESULT
        ' 5) Uses GetIterations
        ' 6) Uses PrintReportInfo
        ' 7) Uses GetNormalizedStatus
        ' 8) Uses Timestamp
        ' -------------------------------------------------
        
        Const C_STEPS_DATASHEET = "Steps"
        Dim iTestStatus, iStepStatus, iIterationStatus
        
        'Statuses at all levels of flow control
        Dim dt, rowcount 'Datasheet with the steps list
        Dim bExitAction, bExitTest, bRun, iStep, iter, oAction, sActionName 
        'For the steps and iterations flow control
        
        Dim arrIterations 'To support iterations
        Dim sFolder, sDatasheet 'For datasheet import
        ' ------------------------------------------------
        ' -------------------------------------------------
        '--- Get the name of the folder from which to import datasheets (same as test)
        sFolder = Environment("TestName")
        '--- Add sheet
        DataTable.AddSheet(C_STEPS_DATASHEET)
        '--- Import steps datasheet
        Call DataTable.ImportSheet(strTestSetsPathName & "\" &  C_STR_TEST_DATA_XLS, C_STEPS_DATASHEET, C_STEPS_DATASHEET)
                
        Set iTestStatus = [As Num](0)
        Set dt = DataTable.GetSheet(C_STEPS_DATASHEET)
        rowcount = dt.GetRowCount
        bExitTest = False
        PrintReportInfo "Test " , "Started at " & Timestamp()
        
        '--- Loop on all steps defined in the datasheet
        For iStep = 1 To rowcount
            bExitAction = False
            dt.SetCurrentRow(iStep)
            sActionName = dt.GetParameter("ACTION_NAME").Value
            bRun = dt.GetParameter("RUN").Value
            
            '--- Check if the step is planned to be executed
            If CStr(bRun) = "TRUE" Then
                '--- Get an instance of the sActionName class
                ASSERT_RESULT(GetClassInstance(oAction, "[" & sActionName & "]"))
                '--- Reset Step status
                Set iStepStatus = [As Num](0)
                '--- Assign Step id
                oAction.StepNum = dt.GetParameter("STEP_ID").Value
                oAction.value = dt.GetParameter("VALUE").Value
                oAction.store_result_as = dt.GetParameter("STORE_RESULT_AS").Value
                '--- Get datasheet name to import (for data-driven actions)
                sDatasheet = dt.GetParameter("DATASHEET").Value
                If Trim(sDatasheet) = "" Then
                    sDatasheet = sActionName
                End If
                
                '--- Check if the Action is data-driven
                If sDatasheet <> "N/A" Then
                    '--- Import datasheet to local
                    Call DataTable.ImportSheet(strTestSetsPathName & "\" & sFolder &"\" & C_STR_TEST_DATA_XLS, sDatasheet, Environment("ActionName"))
                    '--- Assign the new sheet to the step
                    Set oAction.dt = DataTable.LocalSheet
                End If
                
                '--- Get list of iterations (e.g., "1-3,7,13-17") as System.Collections.ArrayList and sort
                Set arrIterations = GetIterations(dt.GetParameter("ITERATIONS").Value)
                If arrIterations.count = 0 Then arrIterations.add "1"
                
                arrIterations.Sort()
                '--- Reset iterations status
                Set iIterationStatus = [As Num](0)
                '--- Send start Step to the log
                PrintReportInfo "Step " & oAction.StepNum & " - Action '" & sActionName & "'", "Started at " & Timestamp()
                
                '--- Loop for each iteration
                For Each iter In arrIterations
                    PrintReportInfo "Step " & oAction.StepNum & " – Action '" &  sActionName & "’", "Started iteration " & iter & " at " & Timestamp()
                    '--- Check if the Action is data-driven
                    If sDatasheet <> "N/A" Then
                        '--- Set the row that corresponds to the current iteration
                        oAction.dt.SetCurrentRow(iter)
                    End If
                    '--- Set the Iteration field of the Action
                    oAction.Iteration = iter
                    ' -------------------------------------
                    '--- Execute the Action
                    ' -------------------------------------
                    On Error Resume Next '--- Try
                    oAction.Run
                    ' -------------------------------------
                    If Err.Number <> 0 Then 'Catch
                        me.ErrorHandler.RunMappedProcedure(Err.Number)
                    End If
                    On Error Goto 0
                    ' -------------------------------------
                    ' --- Get result
                    'On Error Resume Next ' --- Try
                    ' Call results.add(oAction.store_result_as, oAction.result)
                    'If Err.Number <> 0 Then
                    '	me.ErrorHandler.RunMappedProcedure(Err.Number)
                    'End If
                    'On Error Goto 0
                    '--------------------------------------
                    '--- Get the Action status
                    iIterationStatus.[+=]oAction.Status
                    '--- Send iteration result to the log
                    PrintReportInfo "Step " & oAction.StepNum & " - Action '" & sActionName & "'", "Ended iteration " & iter & " at " & Timestamp() & " with status " & GetNormalizedStatus(iIterationStatus)
                    '--- Check the status of the iteration
                    If GetNormalizedStatus(iIterationStatus) > 0 Then
                        '--- Evaluate if a failure condition occurred
                        Eval("b" & dt.GetParameter("ON_FAILURE") & "=TRUE")
                        '--- Check the Exit flags
                        If bExitAction Then Exit For
                        If bExitTest Then Exit For
                    End If
                Next '--- Iteration
                '--- Update the Step status with the iteration status
                iStepStatus.[+=]iIterationStatus
                '--- Send Action result (end) to the log
                PrintReportInfo "Step " & oAction.StepNum & " - Action '" & sActionName & "'", "Ended at " & Timestamp() & " with status " & GetNormalizedStatus(iStepStatus)
                '--- Dispose of the oAction object
                Set oAction = Nothing
            ElseIf CStr(bRun) = "FALSE" Then
                '--- Send skip Step to the log
                PrintReportInfo "Step " & dt.GetParameter("STEP_ID").Value & " - Action '" & sActionName & "'", "Not planned to run"
            Else
                '--- Send no directive for Step to the log
                PrintReportInfo "Step " & dt.GetParameter("STEP_ID").Value & " -Action '" & sActionName & "'","Undefined"
            End If
            '--- Update the Test status with the iteration status
            iTestStatus.[+=]GetNormalizedStatus(iStepStatus)
            '--- Check the Exit flag
            If bExitTest Then Exit For
        Next '--- Step (Action)
        '--- Send Test result (end) to the log
        PrintReportInfo "Test " & Environment("TestName"),"Ended at " & Timestamp() & " with status " & GetNormalizedStatus(iTestStatus)
        '--- Return status
        Run = GetNormalizedStatus(iTestStatus)
    End Function
    ' -----------------------------------------------------
    ' End: Run
    ' -----------------------------------------------------
End Class

Function RunTest()
    Dim oController
    ASSERT_RESULT(GetClassInstance(oController, "Controller"))
    RunTest = oController.Run(Environment("DATA_FOLDER"))
End Function







Function ASSERT_RESULT(ByVal iResult)
    ' --------------------------------------------------------
    ' Function : ASSERT_RESULT
    ' Purpose : Checks if the result triggers a predefined action
    ' Args : ByVal iResult
    ' Returns : The value of iResult (unless the runsession is terminated)
    ' --------------------------------------------------------
    ASSERT_RESULT = CLng(iResult)
    If CLng(iResult) <> CLng(micPass) Then
        Reporter.ReportEvent micWarning, "ASSERT_RESULT", "The action stopped by ASSERT_RESULT"
        Execute(Environment("ON_FAILURE") & "(" & CStr(CLng(iResult)) & ")")
    End If
End Function

Function InfoClassInstance(ByVal p, ByVal msg)
    '---------------------------------------------------------
    'Description: Prints a log message relating to an object
    'Arguments :
    ' p - a reference to the instance
    ' msg - a string
    'Usage : For example, in a Sub Class_Initialize within a Class
    ' InfoClassInstance(me, "Loaded successfully")
    'Changes Log:
    '---------------------------------------------------------
    Print C_OBJ_OF_CLASS_MSG & TypeName(p) & msg & " at " & Timestamp()
End Function

Function GetClassInstance(oInst, ByVal sClass)
    ' --------------------------------------------------------
    ' Function : GetClassInstance
    ' Purpose : Returns an instance of the specified Class
    ' Args : byRef oInst (output variable to return the instance)
    ' ByVal sClass (name of requested Class)
    ' Returns : 0 (success), 1 (failure)
    ' --------------------------------------------------------
    GetClassInstance = 0
    On Error Resume Next
    Execute "Set oInst = new " & sClass
    If Err.Number <> 0 Then
        Set oInst = Nothing
        GetClassInstance = 1
        PrintReportInfo "GetClassInstance","Failed to create an instance of '" & sClass &"'"
        reporter.ReportEvent micFail, "GetClassInstance","Failed to create an instance of '" & sClass &"'"
    End If
End Function

Function GetIterations(ByVal sIterations)
    ' ------------------------------------------------------------------
    ' Function : GetIterations
    ' Purpose : Get array with list of iterations
    ' Args : ByVal sIterations - A comma and hyphen separated
    ' string list with numbers of iterations to be run
    ' Returns : A System.Collections.ArrayList
    ' ------------------------------------------------------------------
    ' Usage : Set DotNetArray = GetIterations("1,3,7-9,15-22")
    ' Print DotNetArray.Count
    ' For each item in DotNetArray
    ' Print item
    ' Next
    ' ------------------------------------------------------------------
    Dim arrRange, min, max, i, j
    Dim arrIterations : Set arrIterations = CreateObject("System.Collections.ArrayList")
    Dim arrTmp : arrTmp = Split(sIterations,",")
    'Parse array with iterations
    For i = 0 To UBound(arrTmp)
        arrRange = Split(arrTmp(i), "-")
        If UBound(arrRange) = 1 Then '--- If is a Range
            min = arrRange(0)
            max = arrRange(1)
            If min > max Then
                Call SwapArgs(min, max)
            End If
            For j = min To max
                arrIterations.Add j
            Next
        Else '--- A single numeric value
            arrIterations.Add arrTmp(i)
        End If
        '--- Dispose of temporary range array
        Erase arrRange
    Next
    '--- Dispose of temporary array
    Erase arrTmp
    '--- Return DotNet array
    Set GetIterations = arrIterations
End Function
' --------------------------------------------------
Function PadNumber(iNum, ByVal iMax)
    ' --------------------------------------------------------
    ' Function : PadNumber
    ' Purpose : Pad a number with zeroes
    ' Args : ByRef iNum (the number to be padded)
    ' ByVal iMax (the max value of the range)
    ' Usage : PadNumber(3, 100) will return the string "003"
    ' Returns : String
    ' --------------------------------------------------------
    'Validates the arguments - If invalid Then it returns the value as is
    If (Not IsNumeric(iNum) Or Not IsNumeric(iMax)) Then
        PadNumber = iNum
        Exit Function
    End If
    If (Abs(iNum) >= Abs(iMax)) Then
        PadNumber = iNum
        Exit Function
    End If
    PadNumber = String(Len(CStr(Abs(iMax))) - Len(CStr(Abs(iNum))), "0") & CStr(Abs(iNum))
End Function

Function Timestamp()
    ' --------------------------------------------------------
    ' Function : Timestamp
    ' Purpose : Build a timestamp string
    ' Args : N/A
    ' Returns : String
    ' --------------------------------------------------------
    Dim sDate, sTime
    sDate=Date()
    sTime=Time()
    Timestamp = Year(sDate) & _
    PadNumber(Month(sDate), 12) & _
    PadNumber(Day(sDate), 31) & "_" & _
    PadNumber(Hour(sTime),24) & _
    PadNumber(Minute(sTime), 60) & _
    PadNumber(Second(sTime), 60)
End Function

Class CNum
    Private m_value
    
    Public Function [=](n)
        value = n
    End Function
    
    Public Function [++]()
        value = value+1
        [++] = value
    End Function
    
    Public Function [--]()
        value = value-1
        [--] = value
    End Function
    
    Public Function [+=](n)
        value = value+n
        [+=] = value
    End Function
    
    Public Function [-=](n)
        value = value-n
        [-=] = value
    End Function
    
    Public Function [*=](n)
        value = value*n
        [*=] = value
    End Function
    
    Public Function [/=](n)
        value = value/n
        [/=] = value
    End Function
    
    Public Function [\=](n)
        value = value\n
        [\=] = value
    End Function
    
    Public default Property Get Value()
    	Value = m_value
    End Property
    
    Public Property Let Value(n)
    	m_value = n
    End Property
    
    Sub Class_Initialize()
        value = 0
    End Sub
End Class

Function [As Num](n)
    Set [As Num] = New CNum
    If IsNumeric(n) Then [As Num].Value = n
End Function

Function [++](n)
    Dim i
    Set i = [As Num](n)
    i.value = n
    i.[++]
    [++] = i
End Function

Function[--](n)
    Dim i
    Set i = [As Num](n)
    i.value = n
    i.[--]
    [--] = i
End Function
