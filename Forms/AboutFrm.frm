VERSION 5.00
Begin VB.Form AboutFrm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "关于 VB6 Helper"
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6780
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
   ScaleHeight     =   1170
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Label LabelX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Microsoft Docs"
      Height          =   255
      Index           =   3
      Left            =   4800
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label LabelX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "参考资料"
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   2
      Left            =   4800
      TabIndex        =   3
      Top             =   120
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
