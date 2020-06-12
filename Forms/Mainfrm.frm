VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Mainfrm 
   BorderStyle     =   0  'None
   Caption         =   "VB6 Helper"
   ClientHeight    =   7365
   ClientLeft      =   150
   ClientTop       =   195
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   13830
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox TitlePicture 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   4080
      ScaleHeight     =   735
      ScaleWidth      =   6975
      TabIndex        =   5
      Top             =   0
      Width           =   6975
      Begin VB6_Helper.NockButton FeedbackButton 
         Height          =   375
         Left            =   4560
         TabIndex        =   7
         ToolTipText     =   "访问 GitHub 本项目进行文档报错反馈"
         Top             =   120
         Width           =   1020
         _extentx        =   1799
         _extenty        =   661
         picture_normal  =   "Mainfrm.frx":038A
         picture_down    =   "Mainfrm.frx":03A6
         picture_hover   =   "Mainfrm.frx":03C2
         stretch         =   0   'False
         caption         =   "文档报错"
         background      =   15725042
         backcolornormal =   15725042
         backcolorhover  =   14737632
         backcolordown   =   16777215
         bordercolornormal=   15725042
         bordercolorhover=   15725042
         bordercolordown =   15725042
         fontsize        =   10
         font            =   "Mainfrm.frx":03E0
         bordercustom    =   11776947
         forecolornormal =   4210752
         forecolorhover  =   4210752
         forecolordown   =   4210752
         text_visible    =   -1  'True
      End
      Begin VB.Label TitleLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "TitleLabel"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   26.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00CC7A00&
         Height          =   690
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   2535
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
      TabIndex        =   2
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
         TabIndex        =   4
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
         TabIndex        =   3
         Top             =   120
         Width           =   600
      End
   End
   Begin RichTextLib.RichTextBox MainRichTextBox 
      Height          =   4095
      Left            =   4080
      TabIndex        =   1
      Top             =   720
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   7223
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      MousePointer    =   1
      Appearance      =   0
      TextRTF         =   $"Mainfrm.frx":0408
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
      LabelEdit       =   1
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
End
Attribute VB_Name = "Mainfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'API 引用
Const ME_MINHEIGHT As Integer = 15 * 300 '窗体最小高度
Const ME_MINWIDTH  As Integer = 15 * 500 '窗体最小宽度
Dim ClickStr       As String




Private Sub FeedbackButton_Click()
    Call Link("https://github.com/Moriafly/VB6-Helper/issues")
End Sub

Private Sub FeedbackButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call myHand
End Sub



Private Sub Form_Load()
    
    
    version = App.Major & "." & App.Minor & "." & App.Revision & " beta"
    statustext = MDIfrm.StatusLabel.Caption
    
    
    Me.Caption = "VB6 Helper" & " " & version

    Call myHand_Load
    Call loadtreeview
    
    
    MainRichTextBox.filename = App.Path & "\Source\Visual Basic 6.0.rtf"
    TitleLabel.Caption = "Visual Basic 6.0"
End Sub

Private Sub LeftTreeView_NodeClick(ByVal Node As MSComctlLib.Node)
    On Error GoTo ErrorHandler
        MDIfrm.MDIStatusBar.BackColor = RGB(0, 122, 204)
        MDIfrm.StatusLabel.Caption = "功能性测试"
        
            MainRichTextBox.filename = App.Path & "\Source\" & LeftTreeView.SelectedItem.Key & ".drtf"
            TitleLabel.Caption = LeftTreeView.SelectedItem.Text
        Exit Sub
ErrorHandler:
        Call rtf(1)
End Sub

Private Sub rtf(n As Integer)
    On Error GoTo ErrorHandler
        If n = 1 Then
            MainRichTextBox.filename = App.Path & "\Source\" & LeftTreeView.SelectedItem.Key & ".rtf"
            TitleLabel.Caption = LeftTreeView.SelectedItem.Text
        ElseIf n = 2 Then
            MainRichTextBox.filename = App.Path & "\Source\" & ClickStr & ".rtf"
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



Private Sub MainRichTextBox_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo ErrorHandler

    MDIfrm.MDIStatusBar.BackColor = RGB(0, 122, 204)
    MDIfrm.StatusLabel.Caption = statustext

    MainMenu.Visible = False
    Dim l As Integer, R As Integer, C As Integer, cfu As Integer, cc As Integer, us As Integer, usl As Integer
    C = MainRichTextBox.SelStart
    cfu = MainRichTextBox.SelUnderline
    l = -1
    cc = Len(MainRichTextBox.Text) - 1
    If cfu <> 0 Then
        us = MainRichTextBox.SelStart
        usl = MainRichTextBox.SelLength
        Dim i As Integer
        For i = C To 0 Step -1
            MainRichTextBox.SelStart = i
            If MainRichTextBox.SelUnderline = cfu Then
            Else
                l = i + 1
                Exit For
            End If
        Next
        R = cc
        If l <> -1 Then
            For i = C To cc Step 1
                MainRichTextBox.SelStart = i
                If MainRichTextBox.SelUnderline = cfu Then
                Else
                    R = i - 1
                    Exit For
                End If
            Next
        End If
        
        MainRichTextBox.SelStart = l
        MainRichTextBox.SelLength = R - l
        
        
        
        ClickStr = MainRichTextBox.SelText
        'ClickStr = Replace(ClickStr, "<", "")
        'ClickStr = Replace(ClickStr, ">", "")
        ClickStr = Mid(ClickStr, InStr(ClickStr, "<") + 1, InStr(ClickStr, ">") - InStr(ClickStr, "<") - 1)
        
        ClickStr = Replace(ClickStr, "%20", " ")

        Call url(ClickStr, x, y)
        

       
        'MainRichTextBox.FileName = App.Path & "\Source\" & ClickStr & ".rtf"
        
        
        MainRichTextBox.SelStart = us
        MainRichTextBox.SelLength = usl
    End If
    
    
            Exit Sub
ErrorHandler:
        Call ErrorCatch
    Resume Next
End Sub

Public Sub url(str As String, x As Single, y As Single)
    MainMenu.Visible = True
    'MsgBox X & vbCrLf & MainRichTextBox.Width
    If x > MainRichTextBox.Width - 960 Then
        MainMenu.Left = x + MainRichTextBox.Left - 960
    Else
        MainMenu.Left = x + MainRichTextBox.Left + 240
    End If

    If y > MainRichTextBox.Height - MainMenu.Height Then
        MainMenu.Top = y + MainRichTextBox.Top - 960
    Else
        MainMenu.Top = y + MainRichTextBox.Top + 240
    End If
    
    
    MenuLabel.Caption = str
    
    MainMenu.Width = 240 + MenuLabel.Width
    MenuShape.Width = MainMenu.Width

End Sub

Private Sub MenuClick_Click()
        On Error GoTo ErrorHandler
        
        MDIfrm.MDIStatusBar.BackColor = RGB(0, 122, 204)
        MDIfrm.StatusLabel.Caption = statustext
    
            MainRichTextBox.filename = App.Path & "\Source\" & ClickStr & ".drtf"
        Exit Sub
ErrorHandler:
        Call rtf(2)
End Sub

Private Sub MenuClick_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call myHand
End Sub



