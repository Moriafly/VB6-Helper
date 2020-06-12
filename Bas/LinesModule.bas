Attribute VB_Name = "LinesModule"
Option Explicit

' This is used by several functions Some may be pointless but I did not feel like weeding out the ones
' used in this code. You can if you want.
Public Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetWindowLong Lib "USER32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "USER32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetClientRect Lib "USER32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function OffsetRect Lib "USER32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function FillRect Lib "USER32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function DrawText Lib "USER32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

Public Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Public Type POINTAPI
   x As Long
   y As Long
End Type

Public Const SFF_SELECTION = &H8000&
Public Const WM_USER = &H400

Public Const EM_EXSETSEL = (WM_USER + 55)
Public Const EM_EXGETSEL = (WM_USER + 52)
Public Const EM_POSFROMCHAR = &HD6&
Public Const EM_CHARFROMPOS = &HD7&
Public Const EM_EXLINEFROMCHAR = (WM_USER + 54)
Public Const EM_GETTEXTRANGE = (WM_USER + 75)
Public Const EM_STREAMIN = (WM_USER + 73)
Public Const EM_HIDESELECTION = WM_USER + 63

Public Const PS_SOLID = 0
Public Const DT_CALCRECT = &H400
Public Const DT_RIGHT = &H2
Public Const DT_VCENTER = &H4
Public Const DT_SINGLELINE = &H20

Public Const GWL_WNDPROC = (-4)
Private Const WM_VSCROLL = &H115
Public lPrevWndProc As Long

Public Function WindowProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If Msg = WM_VSCROLL Then
        ' This takes the scroll method and scrolls the gutter correctly
        Codefrm.DrawLines Codefrm.picLines
    End If
    WindowProc = CallWindowProc(lPrevWndProc, hWnd, Msg, wParam, ByVal lParam)
End Function

Public Function TranslateColor(ByVal clr As OLE_COLOR, Optional hPal As Long = 0) As Long
    If OleTranslateColor(clr, hPal, TranslateColor) Then
        TranslateColor = -1
    End If
End Function
