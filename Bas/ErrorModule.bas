Attribute VB_Name = "ErrorModule"
Option Explicit

Public version As String
Public statustext As String

Public Sub ErrorCatch()
    Mainfrm.MainStatusBar.BackColor = RGB(202, 81, 0)
    Mainfrm.StatusLabel.Caption = "ดํฮ๓ด๚ย๋ฃบ" & Err.Number & "  หตร๗ฃบ" & Err.Description
    Err.Clear
End Sub
