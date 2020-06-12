VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
'*��ģ�����ƣ�CColor
'*��ģ�鹦�ܣ���������ɫ���й���
'*˵����������˸Ķ������ʸ���һ��(progame@cnnb.net)
'*      ��ҳ��http://www.NewXing.com
'*���ߣ�progame  2002-09-29  13:44:40
'***************************************************************
Option Explicit

Private m_color()   As OLE_COLOR

Private m_iCount    As Integer

Public Sub AddColor(color As OLE_COLOR)
'*����һ���µ���ɫ
    m_iCount = m_iCount + 1
    ReDim Preserve m_color(1 To m_iCount)
    m_color(m_iCount) = color
End Sub

Public Function getColorPos(color As OLE_COLOR) As Integer
'*�õ�һ����ɫ������λ��
Dim i   As Integer
    For i = 1 To m_iCount
        If m_color(i) = color Then
            getColorPos = i
            Exit Function
        End If
    Next i
    '*�Ҳ���������-1
    getColorPos = -1
End Function

Public Function getColorHex(color As OLE_COLOR) As String
'*�õ�һ����ɫ��ʮ�������ַ���
Dim s   As String
    s = Hex(RedColor(color))
    s = String(2 - Len(s), "0") & s
    getColorHex = s
    s = Hex(GreenColor(color))
    s = String(2 - Len(s), "0") & s
    getColorHex = getColorHex & s
    s = Hex(BlueColor(color))
    s = String(2 - Len(s), "0") & s
    getColorHex = getColorHex & s
End Function

'**************************************************************
'*ComStr
'*���ܣ�����������ɫ��rtf��ʽ�еĶ���
'*˵����
'*���������
'*    color      ���ͣ�OLE_COLOR
'*���ز�����
'*
'*���ߣ�progame  ���ڣ�2002-09-29  13:50:55
'**************************************************************
Public Function ComStr() As String
Dim i       As Integer

    For i = 1 To m_iCount
        ComStr = ComStr & "\red" & RedColor(m_color(i)) _
            & "\green" & GreenColor(m_color(i)) _
            & "\blue" & BlueColor(m_color(i)) & ";"
    Next i
End Function

Private Function RedColor(color As OLE_COLOR) As Long
    RedColor = color And 255
End Function

Private Function GreenColor(color As OLE_COLOR) As Long
    GreenColor = (color And 65280) / 256
End Function

Private Function BlueColor(color As OLE_COLOR) As Long
    BlueColor = (color And 16711680) / 65536
End Function

Private Sub Class_Initialize()
'*����Ĭ�ϵ�������ɫ
    m_iCount = 2
    AddColor (vbWhite)
    AddColor (vbBlack)
End Sub