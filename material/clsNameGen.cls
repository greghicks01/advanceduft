VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsNameGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public oDct As Object

Private Sub Class_Initialize()
    Set oDct = CreateObject("scripting.dictionary")
    
    gatherHeader
    
End Sub

Sub gatherHeader()
    ' get default data from sheet
    Set dData = ActiveWorkbook.Sheets("Default Data")
    iRow = 2
    Do
    
        oDct.Add dData.Cells(iRow, 2).Value, iRow - 1
        iRow = iRow + 1
    Loop Until dData.Cells(iRow, 2).Value = ""

End Sub

Property Get ColumnNum(ColumnName As String)
    ColumnNum = oDct.Item(ColumnName)
    
End Property

Private Sub Class_Terminate()
    Set oDct = Nothing
    
End Sub
