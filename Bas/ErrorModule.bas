Attribute VB_Name = "ErrorModule"
Option Explicit

Public version As String
Public statustext As String

Public Sub ErrorCatch()
    Select Case Err.Number
    Case 94 '��Чʹ�� Null
        Exit Sub
    End Select
    
    MDIfrm.MDIStatusBar.BackColor = RGB(202, 81, 0)
    MDIfrm.StatusLabel.Caption = "������룺" & Err.Number & "  ˵����" & Err.Description
    Err.Clear
End Sub
