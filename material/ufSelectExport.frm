VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufSelectExport 
   Caption         =   "Export Type"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3135
   OleObjectBlob   =   "ufSelectExport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufSelectExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bCancelled As Boolean

Private Sub btnCancel_Click()
' Purpose:
' Accepts:
' Returns:
    bCancelled = True
    Me.Hide
End Sub

Private Sub btnOK_Click()
' Purpose:
' Accepts:
' Returns:
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
' Purpose:
' Accepts:
' Returns:
    All.Value = True
    bCancelled = False
End Sub
