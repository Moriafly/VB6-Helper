VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Mainfrm 
   Caption         =   "VB6 Helper"
   ClientHeight    =   7365
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   13830
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Mainfrm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   13830
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin VB.PictureBox MainStatusBar 
      Align           =   2  'Align Bottom
      BackColor       =   &H00CC7A00&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   13830
      TabIndex        =   2
      Top             =   6990
      Width           =   13830
      Begin VB.Label StatusLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "功能性测试"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   60
         Width           =   900
      End
   End
   Begin VB.PictureBox MainMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H00CC7A00&
      ForeColor       =   &H00CC7A00&
      Height          =   855
      Left            =   5040
      ScaleHeight     =   855
      ScaleWidth      =   2895
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      Width           =   2895
      Begin VB.Label MenuClick 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "点击访问"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00CC7A00&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   720
      End
      Begin VB.Shape MenuShape 
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00F0F0F0&
         Height          =   855
         Left            =   0
         Top             =   0
         Width           =   2895
      End
      Begin VB.Label MenuLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "          "
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00CC7A00&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   600
      End
   End
   Begin RichTextLib.RichTextBox MainRichTextBox 
      Height          =   4215
      Left            =   4080
      TabIndex        =   1
      Top             =   0
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   7435
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"Mainfrm.frx":038A
   End
   Begin MSComctlLib.TreeView LeftTreeView 
      Height          =   6015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   10610
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu HelpMenu 
      Caption         =   "帮助(&H)"
      Begin VB.Menu FeedbackMenu 
         Caption         =   "反馈(&F)"
      End
      Begin VB.Menu WebsiteMenu 
         Caption         =   "访问官网"
         Shortcut        =   ^W
      End
      Begin VB.Menu Line1 
         Caption         =   "-"
      End
      Begin VB.Menu AboutMenu 
         Caption         =   "关于 VB6 Helper(&A)"
      End
   End
End
Attribute VB_Name = "Mainfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'API 引用
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function LoadCursor Lib "user32.dll" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32.dll" (ByVal hCursor As Long) As Long

Const IDC_HAND     As Long = 32649
Const ME_MINHEIGHT As Integer = 15 * 300 '窗体最小高度
Const ME_MINWIDTH  As Integer = 15 * 400 '窗体最小宽度
Const WEBSITE      As String = "https://github.com/Moriafly/VB6-Helper" '官方网站

Dim myHand_handle  As Long
Dim ClickStr       As String

Private Sub AboutMenu_Click() '打开关于窗体
    AboutFrm.Show 1
    
End Sub

Private Sub FeedbackMenu_Click() '打开反馈窗体
    Feedbackfrm.Show 1
End Sub

Private Sub Form_Load()
    myHand_handle = LoadCursor(0, IDC_HAND)
    
    version = App.Major & "." & App.Minor & "." & App.Revision & " beta"
    statustext = StatusLabel.Caption
    
    
    Me.Caption = "VB6 Helper" & " " & version

    Call loadtreeview
    
End Sub

Private Sub Form_Resize()
    On Error GoTo ErrorHandler
        
        If Me.Height < ME_MINHEIGHT Then Me.Height = ME_MINHEIGHT
        If Me.Width < ME_MINWIDTH Then Me.Width = ME_MINWIDTH
    
        LeftTreeView.Height = Me.ScaleHeight - LeftTreeView.Top - MainStatusBar.Height
        
        MainRichTextBox.Height = LeftTreeView.Height
        MainRichTextBox.Width = Me.ScaleWidth - MainRichTextBox.Left
        
        'MsgBox LeftTreeView.Left & vbCrLf & MainRichTextBox.Left + MainRichTextBox.Width & vbCrLf & Me.Width
        
    Exit Sub
ErrorHandler:
        Call ErrorCatch
    Resume Next
End Sub

Private Sub LeftTreeView_NodeClick(ByVal Node As MSComctlLib.Node)
    On Error GoTo ErrorHandler
        MainStatusBar.BackColor = RGB(0, 122, 204)
        StatusLabel.Caption = "功能性测试"
    
            MainRichTextBox.FileName = App.Path & "\Source\" & LeftTreeView.SelectedItem.Key & ".drtf"
        Exit Sub
ErrorHandler:
        Call rtf(1)
End Sub

Private Sub rtf(n As Integer)
    On Error GoTo ErrorHandler
        If n = 1 Then
            MainRichTextBox.FileName = App.Path & "\Source\" & LeftTreeView.SelectedItem.Key & ".rtf"
        ElseIf n = 2 Then
            MainRichTextBox.FileName = App.Path & "\Source\" & ClickStr & ".rtf"
        ElseIf n = 3 Then
            ShellExecute Me.hWnd, vbNullString, ClickStr, vbNullString, vbNullString, 1
        End If
        
            MainMenu.Visible = False
        Exit Sub
ErrorHandler:
        If n = 2 Then
            Call rtf(3)
            Exit Sub
        End If
        Call ErrorCatch
    Resume Next
End Sub

Private Sub MainRichTextBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ErrorHandler

    MainStatusBar.BackColor = RGB(0, 122, 204)
    StatusLabel.Caption = statustext

    MainMenu.Visible = False
    Dim L As Integer, R As Integer, C As Integer, cfu As Integer, cc As Integer, us As Integer, usl As Integer
    C = MainRichTextBox.SelStart
    cfu = MainRichTextBox.SelUnderline
    L = -1
    cc = Len(MainRichTextBox.Text) - 1
    If cfu <> 0 Then
        us = MainRichTextBox.SelStart
        usl = MainRichTextBox.SelLength
        Dim i As Integer
        For i = C To 0 Step -1
            MainRichTextBox.SelStart = i
            If MainRichTextBox.SelUnderline = cfu Then
            Else
                L = i + 1
                Exit For
            End If
        Next
        R = cc
        If L <> -1 Then
            For i = C To cc Step 1
                MainRichTextBox.SelStart = i
                If MainRichTextBox.SelUnderline = cfu Then
                Else
                    R = i - 1
                    Exit For
                End If
            Next
        End If
        
        MainRichTextBox.SelStart = L
        MainRichTextBox.SelLength = R - L
        
        
        
        ClickStr = MainRichTextBox.SelText
        'ClickStr = Replace(ClickStr, "<", "")
        'ClickStr = Replace(ClickStr, ">", "")
        ClickStr = Mid(ClickStr, InStr(ClickStr, "<") + 1, InStr(ClickStr, ">") - InStr(ClickStr, "<") - 1)
        
        ClickStr = Replace(ClickStr, "%20", " ")

        Call url(ClickStr, X, Y)
        

       
        'MainRichTextBox.FileName = App.Path & "\Source\" & ClickStr & ".rtf"
        
        
        MainRichTextBox.SelStart = us
        MainRichTextBox.SelLength = usl
    End If
    
    
            Exit Sub
ErrorHandler:
        Call ErrorCatch
    Resume Next
End Sub

Public Sub url(str As String, X As Single, Y As Single)
    MainMenu.Visible = True
    MainMenu.Left = X + MainRichTextBox.Left + 240
    MainMenu.Top = Y + MainRichTextBox.Top + 240
    
    MenuLabel.Caption = str
    
    MainMenu.Width = 240 + MenuLabel.Width
    MenuShape.Width = MainMenu.Width

End Sub

Private Sub MenuClick_Click()
        On Error GoTo ErrorHandler
        
        MainStatusBar.BackColor = RGB(0, 122, 204)
        StatusLabel.Caption = statustext
    
            MainRichTextBox.FileName = App.Path & "\Source\" & ClickStr & ".drtf"
        Exit Sub
ErrorHandler:
        Call rtf(2)
End Sub

Private Sub MenuClick_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If myHand_handle <> 0 Then SetCursor myHand_handle
End Sub

Private Sub WebsiteMenu_Click() '访问官网
On Error GoTo ErrorHandler
    ShellExecute Me.hWnd, vbNullString, WEBSITE, vbNullString, vbNullString, 1
    Exit Sub
ErrorHandler:
    Call ErrorCatch
    Resume Next
End Sub
