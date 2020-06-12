VERSION 5.00
Begin VB.MDIForm MDIfrm 
   Appearance      =   0  'Flat
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "VB6 Helper"
   ClientHeight    =   7425
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   15270
   Icon            =   "MDIfrm.frx":0000
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin VB.PictureBox MDIStatusBar 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00CC7A00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   15270
      TabIndex        =   0
      Top             =   7050
      Width           =   15270
      Begin VB6_Helper.NockButton MainfrmButton 
         Height          =   375
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Picture_Normal  =   "MDIfrm.frx":038A
         Picture_Down    =   "MDIfrm.frx":03A6
         Picture_Hover   =   "MDIfrm.frx":03C2
         Stretch         =   0   'False
         Caption         =   "帮助文档"
         BackGround      =   15725042
         BackColorNormal =   14737632
         BackColorHover  =   15790320
         BackColorDown   =   15790320
         BorderColorNormal=   14737632
         BorderColorHover=   15790320
         BorderColorDown =   15790320
         FontSize        =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderCustom    =   11776947
         Style           =   0
         ForeColorNormal =   4210752
         ForeColorHover  =   4210752
         ForeColorDown   =   4210752
         Text_Visible    =   -1  'True
         StretchToText   =   0   'False
      End
      Begin VB6_Helper.NockButton CodefrmButton 
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Picture_Normal  =   "MDIfrm.frx":03DE
         Picture_Down    =   "MDIfrm.frx":03FA
         Picture_Hover   =   "MDIfrm.frx":0416
         Stretch         =   0   'False
         Caption         =   "代码编辑"
         BackGround      =   15725042
         BackColorNormal =   14737632
         BackColorHover  =   15790320
         BackColorDown   =   15790320
         BorderColorNormal=   14737632
         BorderColorHover=   15790320
         BorderColorDown =   15790320
         FontSize        =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderCustom    =   14737632
         Style           =   0
         ForeColorNormal =   4210752
         ForeColorHover  =   4210752
         ForeColorDown   =   4210752
         Text_Visible    =   -1  'True
         StretchToText   =   0   'False
      End
      Begin VB.Label StatusLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "功能性测试"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2520
         TabIndex        =   1
         Top             =   60
         Width           =   900
      End
   End
   Begin VB.Menu HelpMenu 
      Caption         =   "帮助(&H)"
      Begin VB.Menu FeedbackMenu 
         Caption         =   "反馈(&F)"
      End
      Begin VB.Menu WebMenu 
         Caption         =   "访问官网"
         Shortcut        =   ^W
      End
      Begin VB.Menu Line1 
         Caption         =   "-"
      End
      Begin VB.Menu AboutMenu 
         Caption         =   "关于 VB Helper(&A)"
      End
   End
End
Attribute VB_Name = "MDIfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AboutMenu_Click()
    AboutFrm.Show 1
End Sub

Private Sub CodefrmButton_Click()
    Mainfrm.Hide
    Codefrm.Show
    Codefrm.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    
    MainfrmButton.BackColorNormal = RGB(224, 224, 224)
    MainfrmButton.BorderColorNormal = RGB(224, 224, 224)
    
    CodefrmButton.BackColorNormal = RGB(255, 255, 255)
    CodefrmButton.BorderColorNormal = RGB(255, 255, 255)
    
    MainfrmButton.BackColorHover = MainfrmButton.BackColorNormal
    MainfrmButton.BorderColorHover = MainfrmButton.BackColorNormal
    CodefrmButton.BackColorHover = CodefrmButton.BackColorNormal
    CodefrmButton.BorderColorHover = CodefrmButton.BackColorNormal
End Sub

Private Sub CodefrmButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    myHand
End Sub



Private Sub FeedbackMenu_Click()
    Feedbackfrm.Show 1
End Sub

Private Sub MainfrmButton_Click()
    Codefrm.Hide
    Mainfrm.Show
    Mainfrm.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    
    MainfrmButton.BackColorNormal = RGB(255, 255, 255)
    MainfrmButton.BorderColorNormal = RGB(255, 255, 255)
    
    CodefrmButton.BackColorNormal = RGB(224, 224, 224)
    CodefrmButton.BorderColorNormal = RGB(224, 224, 224)
    
    MainfrmButton.BackColorHover = MainfrmButton.BackColorNormal
    MainfrmButton.BorderColorHover = MainfrmButton.BackColorNormal
    CodefrmButton.BackColorHover = CodefrmButton.BackColorNormal
    CodefrmButton.BorderColorHover = CodefrmButton.BackColorNormal
End Sub

Private Sub MainfrmButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    myHand
End Sub

Private Sub MDIForm_Load()

    
    Call MainfrmButton_Click
    
End Sub

Private Sub MDIForm_Resize()
    On Error GoTo ErrorHandler

        Mainfrm.Move 0, 0, MDIfrm.ScaleWidth, MDIfrm.ScaleHeight
    
        Mainfrm.LeftTreeView.Height = Mainfrm.ScaleHeight - Mainfrm.LeftTreeView.Top
        
        Mainfrm.MainRichTextBox.Height = Mainfrm.LeftTreeView.Height - Mainfrm.MainRichTextBox.Top
        Mainfrm.MainRichTextBox.Width = Mainfrm.ScaleWidth - Mainfrm.MainRichTextBox.Left
        
        Mainfrm.TitlePicture.Width = Mainfrm.MainRichTextBox.Width
        
        'MsgBox LeftTreeView.Left & vbCrLf & MainRichTextBox.Left + MainRichTextBox.Width & vbCrLf & Me.Width
        
        Mainfrm.FeedbackButton.Move Mainfrm.ScaleWidth - Mainfrm.FeedbackButton.Width - Mainfrm.LeftTreeView.Width - (Mainfrm.TitlePicture.Height - Mainfrm.FeedbackButton.Height) / 2, (Mainfrm.TitlePicture.Height - Mainfrm.FeedbackButton.Height) / 2
    
        Codefrm.Move 0, 0, MDIfrm.ScaleWidth, MDIfrm.ScaleHeight
        Codefrm.CodeTextBox.Move 120, 0, Me.ScaleWidth - 120, Me.ScaleHeight
        
    If (WindowState = 0) Then
        If (Me.Width < 6000) Then '限制最小宽度bai
            Me.Enabled = False
            Me.Width = 6000
            Me.Enabled = True
        End If
        If (Me.Height < 4000) Then '限制最小高度
            Me.Enabled = False
            Me.Height = 4000
            Me.Enabled = True
        End If
    End If
        
    Exit Sub
ErrorHandler:
        Call ErrorCatch
    Resume Next
End Sub

Private Sub WebMenu_Click()
    Call Link(WEBSITE)
End Sub
