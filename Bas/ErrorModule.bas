Attribute VB_Name = "ErrorModule"
Option Explicit

Public version As String
Public statustext As String

Public Sub ErrorCatch()
    Select Case Err.Number
    Case 94 '无效使用 Null
        Exit Sub
    End Select
    
    Mainfrm.MainStatusBar.BackColor = RGB(202, 81, 0)
    Mainfrm.StatusLabel.Caption = "错误代码：" & Err.Number & "  说明：" & Err.Description
    Err.Clear
End Sub
