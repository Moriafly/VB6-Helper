VERSION 5.00
Begin VB.Form Feedbackfrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "反馈"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2910
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Feedbackfrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   2910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Label LinkFeedback 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GitHub 在线反馈"
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
      Left            =   360
      TabIndex        =   5
      Top             =   1920
      Width           =   1440
   End
   Begin VB.Label LabelX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "反馈"
      ForeColor       =   &H00CC7A00&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   360
   End
   Begin VB.Label LabelX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开发维护程序"
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   3
      Top             =   840
      Width           =   1080
   End
   Begin VB.Label LabelX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "提供、维护和审核文档"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   1800
   End
   Begin VB.Label LabelX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "加入 VB6 Helper 项目"
      ForeColor       =   &H00CC7A00&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1830
   End
   Begin VB.Label LabelX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "联系 QQ ：1515390445"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   1200
      Width           =   2010
   End
End
Attribute VB_Name = "Feedbackfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub LinkFeedback_Click()
    Call Link("https://github.com/Moriafly/VB6-Helper/issues")
End Sub

Private Sub LinkFeedback_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call myHand
End Sub
