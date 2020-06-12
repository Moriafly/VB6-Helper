VERSION 5.00
Begin VB.Form AboutFrm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "关于 VB6 Helper"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10305
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "AboutFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   10305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   4815
      Left            =   3480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "AboutFrm.frx":000C
      Top             =   1080
      Width           =   6735
   End
   Begin VB.Label LabelX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GNU General Public License v3.0"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   5
      Left            =   4080
      TabIndex        =   7
      Top             =   600
      Width           =   3915
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3480
      Picture         =   "AboutFrm.frx":6D6F
      Top             =   480
      Width           =   480
   End
   Begin VB.Label LabelX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开源说明"
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   4
      Left            =   3480
      TabIndex        =   6
      Top             =   120
      Width           =   720
   End
   Begin VB.Label LabelMicrosoftDocs 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Microsoft Docs"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00CC7A00&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "访问 Microsoft Docs 网站"
      Top             =   1680
      Width           =   1395
   End
   Begin VB.Label LabelX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "参考资料"
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   720
   End
   Begin VB.Label LabelX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright(C) 2020 Dirror"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2160
   End
   Begin VB.Label VerLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "版本："
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   540
   End
   Begin VB.Label LabelX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "程序：虚而遨游"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1260
   End
End
Attribute VB_Name = "AboutFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    VerLabel.Caption = "版本：" & version
    
End Sub

Private Sub LabelMicrosoftDocs_Click()
    Call Link("https://docs.microsoft.com/zh-cn/")
End Sub

Private Sub LabelMicrosoftDocs_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call myHand
End Sub
