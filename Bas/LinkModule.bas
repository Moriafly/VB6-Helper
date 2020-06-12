Attribute VB_Name = "LinkModule"
Option Explicit
'网页
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'手形鼠标
Private Declare Function LoadCursor Lib "user32.dll" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32.dll" (ByVal hCursor As Long) As Long

Public Const WEBSITE      As String = "https://github.com/Moriafly/VB6-Helper" '官方网站
Public Const IDC_HAND     As Long = 32649

Dim myHand_handle  As Long

Public Sub myHand_Load()
    myHand_handle = LoadCursor(0, IDC_HAND)
End Sub

Public Sub myHand()
    If myHand_handle <> 0 Then SetCursor myHand_handle
End Sub

Public Sub Link(str As String)
On Error GoTo ErrorHandler
        ShellExecute Mainfrm.hWnd, vbNullString, str, vbNullString, vbNullString, 1
    Exit Sub
ErrorHandler:
    Call ErrorCatch
    Resume Next
End Sub
