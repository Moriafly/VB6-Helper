Attribute VB_Name = "ErrorModule"
Option Explicit

Public version As String
Public statustext As String

Public Sub ErrorCatch()
    Select Case Err.Number
    Case 94 '��Чʹ�� Null
        Exit Sub
    End Select
    
    Mainfrm.MainStatusBar.BackColor = RGB(202, 81, 0)
    Mainfrm.StatusLabel.Caption = "������룺" & Err.Number & "  ˵����" & Err.Description
    Err.Clear
End Sub
