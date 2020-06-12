VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Codefrm 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15720
   BeginProperty Font 
      Name            =   "Î¢ÈíÑÅºÚ"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   15720
   ShowInTaskbar   =   0   'False
   Begin RichTextLib.RichTextBox CodeTextBox 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   9551
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"Codefrm.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Codefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents Syntax   As CSyntax
Attribute Syntax.VB_VarHelpID = -1
Private Sub CodeTextBox_Change()
    Set Syntax = Nothing
    Set Syntax = New CSyntax
    
    Syntax.ReadFile App.Path & "\Stx\VB6.stx"

    Highlight
End Sub


Private Sub Highlight()

Dim t       As Single
Dim lPos    As Long

    If Syntax Is Nothing Then
        Exit Sub
    End If
    'Screen.MousePointer = vbHourglass
    Screen.MousePointer = vbDefault
    
    lPos = CodeTextBox.SelStart


    t = Timer
    Syntax.HighLightRichEdit CodeTextBox
    'MsgBox Timer - t

    
    On Error Resume Next
    
    CodeTextBox.SelStart = lPos
    'staInfo.Panels(1).Text = "File Length:" & Len(RTXT.Text) & "  Use Seconds:" & Timer - t
    CodeTextBox.SetFocus
    
    'txtInfo.Text = RTXT.TextRTF
    
    
    
    'CodeTextBox.SelStart = 0
    'CodeTextBox.SelLength = Len(CodeTextBox.Text)
    'CodeTextBox.SelFontSize = 20
    'CodeTextBox.SelStart = Len(CodeTextBox.Text)
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub CodeTextBox_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyTab
        
    End Select
End Sub

Private Sub Form_Load()
    Me.Move 0, 0, MDIfrm.ScaleWidth, MDIfrm.ScaleHeight
    
    CodeTextBox.Move 120, 0, Me.ScaleWidth - 120, Me.ScaleHeight
End Sub
