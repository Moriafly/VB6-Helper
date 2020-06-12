VERSION 5.00
Begin VB.UserControl N_Shape 
   BackColor       =   &H00EFF1F2&
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3540
   BeginProperty Font 
      Name            =   "Î¢ÈíÑÅºÚ"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   33
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   236
   ToolboxBitmap   =   "N_Shape.ctx":0000
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1680
      Top             =   0
   End
   Begin VB.Image Image_Button 
      Height          =   435
      Left            =   2280
      Picture         =   "N_Shape.ctx":0312
      Top             =   0
      Visible         =   0   'False
      Width           =   2550
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NShape1"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   705
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   120
      Top             =   120
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00EFF1F2&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00B3B3B3&
      FillColor       =   &H00EFF1F2&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   1185
   End
End
Attribute VB_Name = "N_Shape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'(C) Norksoft
'Dirror ÍêÈ«ÒýÓÃ

Private Const LOGPIXELSY = 90
Private Const LF_FACESIZE = 32
Private Const DT_BOTTOM = &H8
Private Const DT_CENTER = &H1
Private Const DT_LEFT = &H0
Private Const DT_RIGHT = &H2
Private Const DT_TOP = &H0
Private Const DT_VCENTER = &H4
Private Const DT_SINGLELINE = &H20
Private Const DT_WORDBREAK = &H10
Private Const DT_CALCRECT = &H400
Private Const DT_END_ELLIPSIS = &H8000
Private Const DT_MODIFYSTRING = &H10000
Private Const DT_WORD_ELLIPSIS = &H40000

'Cores
Private Const COLOR_BTNHIGHLIGHT = 20
Private Const COLOR_BTNSHADOW = 16
Private Const COLOR_BTNFACE = 15
Private Const COLOR_HIGHLIGHT = 13
Private Const COLOR_ACTIVEBORDER = 10
Private Const COLOR_WINDOWFRAME = 6
Private Const COLOR_3DDKSHADOW = 21
Private Const COLOR_3DLIGHT = 22
Private Const COLOR_INFOTEXT = 23
Private Const COLOR_INFOBK = 24
Private Const PATCOPY = &HF00021
Private Const SRCCOPY = &HCC0020
Private Const PS_SOLID = 0
Private Const PS_DASHDOT = 3
Private Const PS_DASHDOTDOT = 4
Private Const PS_DOT = 2
Private Const PS_DASH = 1
Private Const PS_ENDCAP_FLAT = &H200

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

'Api's utilizadas pelo componente
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Declare Function SelectClipPath Lib "gdi32" (ByVal hdc As Long, ByVal iMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
'Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long ' api repetida
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long

'Para animar o icon
Enum PicStates
    picNothing = 0
    picDown = 1
    picHover = 2
    picNorm = 3
End Enum
Dim PicState As PicStates
Dim m_Picture_Normal As New StdPicture
Dim m_Picture_Hover As New StdPicture
Dim m_Picture_Down As New StdPicture
Dim m_MouseInside As Boolean

'Animar o botao/cor da shape
Dim isOver As Boolean
Dim m_State As Integer
Event Click()
Event DblClick()
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseEnter()
Event MouseLeave()
Public Event MouseOver()
Public Event MouseOut()
Public Event Resize()

'Eventos para as teclas
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)

'Variável para saber se o botao tem icon
Dim botao_com_icon As Boolean
Dim texto_visivel As Boolean
Dim ajustar_ao_texto As Boolean

'Cores utilizadas pelos eventos do botao
Dim cor_fundo_normal As OLE_COLOR
Dim cor_fundo_hover As OLE_COLOR
Dim cor_fundo_down As OLE_COLOR

Dim cor_contorno_normal As OLE_COLOR
Dim cor_contorno_hover As OLE_COLOR
Dim cor_contorno_down As OLE_COLOR
Dim cor_contorno_original As OLE_COLOR
Dim cor_contorno_custom As OLE_COLOR

Dim cor_letra_normal As OLE_COLOR
Dim cor_letra_hover As OLE_COLOR
Dim cor_letra_down As OLE_COLOR

'Variavel para saber se é para alterar a cor do border ao receber o focus
Dim alterar_cor_contorno As Boolean

Private Sub Image1_Click()
    'Evento click
    RaiseEvent Click
End Sub

Private Sub Image1_DblClick()
    'Evento duploclick
    RaiseEvent DblClick
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Evento mousedown
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Evento mousemove
    RaiseEvent MouseMove(Button, Shift, X, Y)

    If Button = vbLeftButton Then
        Exit Sub
    End If
    
    If GetCapture() <> UserControl.hWnd Then
        SetCapture (UserControl.hWnd)
        If Not Image1.Picture = Picture_Hover Then
            Image1.Picture = Picture_Hover
        End If
    Else
        Dim pt As POINTAPI
        pt.X = X
        pt.Y = Y
        ClientToScreen UserControl.hWnd, pt
        If WindowFromPoint(pt.X, pt.Y) <> UserControl.hWnd Then
            Refresh
            If Button <> vbLeftButton Then
                ReleaseCapture
                Image1.Picture = Picture_Normal
                RaiseEvent MouseOut
            End If
            End If
    End If
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Evento mouseup
    RaiseEvent MouseUp(Button, Shift, X, Y)
    RaiseEvent Click
End Sub

Private Sub Label1_Change()
    'Ajustar o botao ao alterar o texto da label
    Call Ajustar_Botao
End Sub

Private Sub Label1_Click()
    'Atalho para
    UserControl_Click
End Sub

Private Sub Label1_DblClick()
    'Atalho para
    UserControl_DblClick
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Evento mousedown
    RaiseEvent MouseDown(Button, Shift, X, Y)
    If Button = 1 Then
        m_State = 2
        Call DrawState
    End If
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Evento mousemove
    RaiseEvent MouseMove(Button, Shift, X, Y)
    If Button < 2 Then
        If Not CheckMouseOver Then
            isOver = False
            m_State = 0
            Call DrawState
        Else
            If Button = 0 And Not isOver Then
                Timer1.Enabled = True
                isOver = True
                RaiseEvent MouseEnter
                m_State = 1
                Call DrawState
            ElseIf Button = 1 Then
                isOver = True
                m_State = 2
                Call DrawState
                isOver = False
            End If
        End If
    End If

    'Icon do botão
    If Button = vbLeftButton Then
        Exit Sub
    End If

    If GetCapture() <> UserControl.hWnd Then
        SetCapture (UserControl.hWnd)
        If Not Image1.Picture = Picture_Hover Then
            Image1.Picture = Picture_Hover
            m_MouseInside = True
        End If
    Else
        Dim pt As POINTAPI
        pt.X = X
        pt.Y = Y
        ClientToScreen UserControl.hWnd, pt
        If WindowFromPoint(pt.X, pt.Y) <> UserControl.hWnd Then
            Refresh
            If Button <> vbLeftButton Then
                ReleaseCapture
                Image1.Picture = Picture_Normal
                m_MouseInside = False
                RaiseEvent MouseOut
            End If
            End If
    End If
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Evento mouseup
    RaiseEvent MouseUp(Button, Shift, X, Y)
    If CheckMouseOver Then
        m_State = 1
    Else
        m_State = 0
    End If
    
    RaiseEvent Click
    Call DrawState
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    'Evento keypress
    LastButton = 1
    Call UserControl_Click
End Sub

Private Sub UserControl_Click()
    'Evento click
    'RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    'Evento duploclick
    RaiseEvent DblClick
End Sub

Private Sub UserControl_GotFocus()
    'Alterar a cor do border
    If BorderGotFocus = True Then
        Shape1.BorderColor = cor_contorno_custom
    End If
End Sub

Private Sub UserControl_Initialize()
    'Iniciando o componente
    If Not Picture_Normal Is Nothing Then
        Set Image1.Picture = Picture_Normal
    End If

    'Call Ajustar_Botao
End Sub

Public Property Get Caption() As String
    'Escolher o nome do botão
    Caption = Label1.Caption
End Property

Public Property Let Caption(New_Value As String)
    'Alterar o caption para o novo texto
    Label1.Caption = New_Value
    PropertyChanged "Caption"
End Property

Private Sub UserControl_InitProperties()
    'Ler as propriedades do botão
    On Error GoTo Corrige_UserControl_InitProperties
    Caption = UserControl.Extender.Name '"NShape1"
    Enabled = True
    BackGround = &HEFF1F2
    BackColorNormal = &HEFF1F2
    BackColorHover = &HE0E0E0
    BackColorDown = RGB(255, 255, 255)
    BorderColorNormal = &HB3B3B3
    BorderColorHover = &HB3B3B3
    BorderColorDown = &HB3B3B3
    FontSize = 8
    FontBold = False
    Font = "Verdana" '"MS Sans Serif"
    BorderGotFocus = False
    BorderCustom = &HB3B3B3
    ForeColorNormal = RGB(0, 0, 0) '&H0&
    ForeColorHover = RGB(0, 0, 0)
    ForeColorDown = RGB(0, 0, 0)
    Set Image1.Picture = Picture_Normal
    Call Ajustar_Botao
    Graphical = False
    Text_Visible = True
    StretchToText = True
    
Corrige_UserControl_InitProperties:
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    'Evento keydown, atalho para as teclas
    Select Case KeyCode
        Case vbKeyRight, vbKeyDown
            Call SendKeys("{TAB}")
            Case vbKeyLeft, vbKeyUp
            Call SendKeys("+{TAB}")
        
        Case vbKeyReturn
            UserControl_Click
    End Select
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    'Evento keypress
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    'Evento keyup
    RaiseEvent KeyUp(KeyCode, Shift)
    UserControl.Refresh
End Sub

Private Sub UserControl_LostFocus()
    'Repor a cor original do border
    If BorderGotFocus = True Then
        Shape1.BorderColor = cor_contorno_original
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Evento mousedown
    Image1.Picture = Picture_Down
    RaiseEvent MouseDown(Button, Shift, X, Y)
    If Button = 1 Then
        m_State = 2
        Call DrawState
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Evento mousemove
    RaiseEvent MouseMove(Button, Shift, X, Y)
    If Button < 2 Then
        If Not CheckMouseOver Then
            isOver = False
            m_State = 0
            Call DrawState
        Else
            If Button = 0 And Not isOver Then
                Timer1.Enabled = True
                isOver = True
                RaiseEvent MouseEnter
                m_State = 1
                Call DrawState
            ElseIf Button = 1 Then
                isOver = True
                m_State = 2
                Call DrawState
                isOver = False
            End If
        End If
    End If

    'Icon do botão
    If Button = vbLeftButton Then
        Exit Sub
    End If

    If GetCapture() <> UserControl.hWnd Then
        SetCapture (UserControl.hWnd)
        If Not Image1.Picture = Picture_Hover Then
            Image1.Picture = Picture_Hover
            m_MouseInside = True
        End If
    Else
        Dim pt As POINTAPI
        pt.X = X
        pt.Y = Y
        ClientToScreen UserControl.hWnd, pt
        If WindowFromPoint(pt.X, pt.Y) <> UserControl.hWnd Then
            Refresh
            If Button <> vbLeftButton Then
                ReleaseCapture
                Image1.Picture = Picture_Normal
                m_MouseInside = False
                RaiseEvent MouseOut
            End If
            End If
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Evento mouseup
    Image1.Picture = Picture_Hover
    RaiseEvent MouseUp(Button, Shift, X, Y)
    
    If CheckMouseOver Then
        m_State = 1
    Else
        m_State = 0
    End If
    
    RaiseEvent Click
    Call DrawState
End Sub

Private Function CheckMouseOver() As Boolean
    'Efectuar o over do botao
    Dim pt As POINTAPI
    GetCursorPos pt
    CheckMouseOver = (WindowFromPoint(pt.X, pt.Y) = UserControl.hWnd)
End Function

Private Sub DrawState()
    On Error Resume Next
    If m_State = 1 Then 'mouse hover
        Shape1.FillColor = cor_fundo_hover
        Label1.ForeColor = cor_letra_hover
        Shape1.BorderColor = cor_contorno_hover
        'If Graphical = True Then Image1.Picture = Picture_Hover
    
    ElseIf m_State = 2 Then 'mouse down
        Shape1.FillColor = cor_fundo_down
        Label1.ForeColor = cor_letra_down
        Shape1.BorderColor = cor_contorno_down
        'If Graphical = True Then Image1.Picture = Picture_Down
        
    Else 'normal
        Shape1.FillColor = cor_fundo_normal
        Label1.ForeColor = cor_letra_normal
        Shape1.BorderColor = cor_contorno_normal
        'If Graphical = True Then Image1.Picture = Picture_Normal
    End If
End Sub

Private Sub Timer1_Timer()
    'Animar o botão
    If Not CheckMouseOver Then
        Timer1.Enabled = False
        isOver = False
        RaiseEvent MouseLeave
        m_State = 0
        Call DrawState
    End If
End Sub

Public Property Get Style() As ShapeConstants
    'Escolher o novo estilo do botão
    Style = Shape1.Shape
End Property

Public Property Let Style(New_Value As ShapeConstants)
    'Alterar o estilo do botão
    Shape1.Shape = New_Value
    'shpShadow.Shape = new_Value
    PropertyChanged "Style"
End Property

Public Property Get FontSize() As Integer
    'Escolher o tamanho da letra
    FontSize = Label1.FontSize
End Property

Public Property Let FontSize(New_Value As Integer)
    'Alterar o tamanho da letra
    Label1.FontSize = New_Value
    PropertyChanged "FontSize"
End Property

Public Property Get FontBold() As Boolean
    'Indicar se que ou não a letra em negrito
    FontBold = Label1.FontBold
End Property

Public Property Let FontBold(New_Value As Boolean)
    'Alterar a letra para negrito se for o caso
    Label1.FontBold = New_Value
    PropertyChanged "FontBold"
End Property

Public Property Get Enabled() As Boolean
    'Escolher se o botão fica activo ou não
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(New_Value As Boolean)
    'Verificar se o botão fica activo ou não
    UserControl.Enabled = New_Value
    PropertyChanged "Enabled"
    
    If New_Value = True Then
        Label1.ForeColor = cor_letra_normal
    Else
        Label1.ForeColor = &HB3B3B3  'Cinzento escuro
    End If
    UserControl.Refresh
End Property

Public Property Get BorderGotFocus() As Boolean
    'Escolher se o botão fica activo ou não
    BorderGotFocus = alterar_cor_contorno
End Property

Public Property Let BorderGotFocus(New_Value As Boolean)
    'Verificar se o botão fica activo ou não
    alterar_cor_contorno = New_Value
    PropertyChanged "BorderGotFocus"
End Property

Public Property Get Picture_Normal() As StdPicture
    'Obter a imagem normal
    Set Picture_Normal = m_Picture_Normal
End Property

Public Property Set Picture_Normal(vNewPic As StdPicture)
    'Escolher a imagem normal
    Set m_Picture_Normal = vNewPic
    PropertyChanged "Picture_Normal"
    Image1.Picture = Picture_Normal
End Property

Public Property Get Picture_Down() As StdPicture
    'Obter a imagem down
    Set Picture_Down = m_Picture_Down
End Property

Public Property Set Picture_Down(vNewPic As StdPicture)
    'Escolher a imagem down
    Set m_Picture_Down = vNewPic
    PropertyChanged "Picture_Down"
End Property

Public Property Get Picture_Hover() As StdPicture
    'Obter a imagem over
    Set Picture_Hover = m_Picture_Hover
End Property

Public Property Set Picture_Hover(vNewPic As StdPicture)
    'Escolher a imagem over
    Set m_Picture_Hover = vNewPic
    PropertyChanged "Picture_Hover"
End Property

Private Sub UserControl_Paint()
    'Escolher a imagem normal para o estado normal do botao
    Image1.Picture = Picture_Normal
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    'Ler as propriedades do control
    On Error GoTo Corrige_UserControl_ReadProperties
    With PropBag
        Set m_Picture_Normal = .ReadProperty("Picture_Normal", Nothing)
        Set m_Picture_Down = .ReadProperty("Picture_Down", Nothing)
        Set m_Picture_Hover = .ReadProperty("Picture_Hover", Nothing)
        Image1.Stretch = .ReadProperty("Stretch", False)
        
        Caption = .ReadProperty("Caption", "NShape1")
        BackGround = .ReadProperty("BackGround", "&HFF8080")
        BackColorNormal = .ReadProperty("BackColorNormal", "&H8080FF")
        BackColorHover = .ReadProperty("BackColorHover", "&H80C0FF")
        BackColorDown = .ReadProperty("BackColorDown", "&H80FFFF")
        BorderColorNormal = .ReadProperty("BorderColorNormal", "&H00FFFF80&")
        BorderColorHover = .ReadProperty("BorderColorHover", "&H00FFFF80&")
        BorderColorDown = .ReadProperty("BorderColorDown", "&H00FFFF80&")
        FontSize = .ReadProperty("FontSize", "8")
        FontBold = .ReadProperty("FontBold", "False")
        Font = .ReadProperty("Font", "MS Sans Serif")
        Enabled = .ReadProperty("Enabled", "True")
        BorderGotFocus = .ReadProperty("BorderGotFocus", "False")
        BorderCustom = .ReadProperty("BorderCustom", "&H80FF80")
        Style = .ReadProperty("Style", "4")
        Graphical = .ReadProperty("Graphical", "False")
        ForeColorNormal = .ReadProperty("ForeColorNormal", "&H8080FF")
        ForeColorHover = .ReadProperty("ForeColorHover", "&H8080FF")
        ForeColorDown = .ReadProperty("ForeColorDown", "&H8080FF")
        novo_icon = .ReadProperty("List_Icons", 1)
        Text_Visible = .ReadProperty("Text_Visible", "False")
        StretchToText = .ReadProperty("StretchToText", "True")
    End With
    
Corrige_UserControl_ReadProperties:
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    'Escrever as propriedades do control
    On Error GoTo Corrige_UserControl_WriteProperties
    With PropBag
        .WriteProperty "Picture_Normal", m_Picture_Normal, 0
        .WriteProperty "Picture_Down", m_Picture_Down, 0
        .WriteProperty "Picture_Hover", m_Picture_Hover, 0
        .WriteProperty "Stretch", Image1.Stretch
        
        Call .WriteProperty("Caption", Caption, "")
        Call .WriteProperty("BackGround", BackGround, "&HFF8080")
        Call .WriteProperty("BackColorNormal", BackColorNormal, "&H8080FF")
        Call .WriteProperty("BackColorHover", BackColorHover, "&H80C0FF")
        Call .WriteProperty("BackColorDown", BackColorDown, "&H80FFFF")
        Call .WriteProperty("BorderColorNormal", BorderColorNormal, "&H00FFFF80&")
        Call .WriteProperty("BorderColorHover", BorderColorHover, "&H00FFFF80&")
        Call .WriteProperty("BorderColorDown", BorderColorDown, "&H00FFFF80&")
        Call .WriteProperty("FontSize", FontSize, "8")
        Call .WriteProperty("FontBold", FontBold, "False")
        Call .WriteProperty("Font", Font, "MS Sans Serif")
        Call .WriteProperty("Enabled", Enabled, "True")
        Call .WriteProperty("BorderGotFocus", BorderGotFocus, "False")
        Call .WriteProperty("BorderCustom", BorderCustom, "&H80FF80")
        Call .WriteProperty("Style", Style, "4")
        Call .WriteProperty("Graphical", Graphical, "False")
        Call .WriteProperty("ForeColorNormal", ForeColorNormal, "&HFF8080")
        Call .WriteProperty("ForeColorHover", ForeColorHover, "&HFF8080")
        Call .WriteProperty("ForeColorDown", ForeColorDown, "&HFF8080")
        Call .WriteProperty("List_Icons", novo_icon, 1)
        Call .WriteProperty("Text_Visible", Text_Visible, "False")
        Call .WriteProperty("StretchToText", StretchToText, "True")
    End With
    
Corrige_UserControl_WriteProperties:
End Sub

Public Property Get Graphical() As Boolean
    'Escolher se o botão fica activo ou não
    Graphical = botao_com_icon
End Property

Public Property Let Graphical(New_Value As Boolean)
    'Verificar se o botão fica activo ou não
    botao_com_icon = New_Value
    PropertyChanged "Graphical"

    If botao_com_icon = True Then
        Image1.Visible = True
    Else
        Image1.Visible = False
    End If
    
    Call Ajustar_Botao
End Property

Public Property Get StretchToText() As Boolean
    'Escolher se o botão fica activo ou não
    StretchToText = ajustar_ao_texto
End Property

Public Property Let StretchToText(New_Value As Boolean)
    'Verificar se o botão fica activo ou não
    ajustar_ao_texto = New_Value
    PropertyChanged "StretchToText"

    Call Ajustar_Botao
End Property

Public Property Get Text_Visible() As Boolean
    'Escolher se o botão fica activo ou não
    Text_Visible = texto_visivel
End Property

Public Property Let Text_Visible(New_Value As Boolean)
    'Verificar se o botão fica activo ou não
    texto_visivel = New_Value
    PropertyChanged "Text_Visible"

    If texto_visivel = True Then
        Label1.Visible = True
    Else
        Label1.Visible = False
    End If
    
    Call Ajustar_Botao
End Property

Private Sub UserControl_Resize()
    'Desenhar o botão ajustando os controles do form
    RaiseEvent Resize
    Ajustar_Botao
End Sub

Public Property Get Stretch() As Boolean
    'Ajustar a imagem do botão
    Stretch = Image1.Stretch
End Property

Public Property Let Stretch(vNewValue As Boolean)
    'Verificar se a imagem é para estar ajustada ou não
    Image1.Stretch = vNewValue
    PropertyChanged "Stretch"
End Property

Public Property Get CurPicture() As PicStates
    'Verificar qual é a imagem a ser vista
    If Image1.Picture = 0 Then
        CurPicture = picNothing
    ElseIf Image1.Picture = Picture_Normal Then
        CurPicture = picNorm
    ElseIf Image1.Picture = Picture_Down Then
        CurPicture = picDown
    ElseIf Image1.Picture = Picture_Hover Then
        CurPicture = picHover
    End If
End Property

Public Property Get MouseInside() As Boolean
    'Mouse em cima do cntrol
    MouseInside = m_MouseInside
End Property

Public Sub Ajustar_Botao()
    'Actualizar o botão
    On Error Resume Next
    With UserControl
        If ajustar_ao_texto = True Then
            If texto_visivel = True Then
                If botao_com_icon = False Then
                    .Width = Screen.TwipsPerPixelX * (Label1.Width + 16)
                Else
                    .Width = Screen.TwipsPerPixelX * ((Label1.Width + 16) + Image1.Width + 8)
                End If
            Else
                If botao_com_icon = False Then
                    .Width = Screen.TwipsPerPixelX * (16)
                Else
                    .Width = Screen.TwipsPerPixelX * (Image1.Width + 16)
                End If
            End If
            
        Else
            '.Width = Screen.TwipsPerPixelX * Image_Button.Width
        End If
    End With
    
    With Shape1
        .Top = 0
        .Height = UserControl.ScaleHeight
        .Left = 0
        .Width = UserControl.ScaleWidth
    End With
    
    With Image1
        .Top = (UserControl.ScaleHeight - .Height) / 2
        .Left = 8
    End With
    
    With Label1
        .Top = (UserControl.ScaleHeight - .Height) / 2
        If ajustar_ao_texto = True Then
            If botao_com_icon = False Then
                .Left = 8
            Else
                .Left = Image1.Left + Image1.Width + 8
            End If
        
        Else
            .Left = (UserControl.ScaleWidth - .Width) / 2 '8
        End If
    End With
End Sub

Public Property Get ForeColorNormal() As OLE_COLOR
    'Escolher a cor inicial de fundo do botão
    ForeColorNormal = Label1.ForeColor
End Property

Public Property Let ForeColorNormal(New_Value As OLE_COLOR)
    'Alterar a cor inicial de fundo do botão
    cor_letra_normal = New_Value
    Label1.ForeColor = New_Value
    PropertyChanged "ForeColorNormal"
End Property

Public Property Get ForeColorHover() As OLE_COLOR
    'Escolher a cor inicial da letra do botão
    ForeColorHover = cor_letra_hover
End Property

Public Property Let ForeColorHover(New_Value As OLE_COLOR)
    'Alterar a cor inicial da letra do botão
    cor_letra_hover = New_Value
    'Shape1.FillColor = new_Value
    PropertyChanged "ForeColorHover"
End Property

Public Property Get ForeColorDown() As OLE_COLOR
    'Escolher a cor inicial da letra do botão
    ForeColorDown = cor_letra_down
End Property

Public Property Let ForeColorDown(New_Value As OLE_COLOR)
    'Alterar a cor inicial da letra do botão
    cor_letra_down = New_Value
    'Shape1.FillColor = new_Value
    PropertyChanged "ForeColorDown"
End Property

Public Property Get BackGround() As OLE_COLOR
    'Escolher a cor inicial de fundo do botão
    BackGround = UserControl.BackColor
End Property

Public Property Let BackGround(New_Value As OLE_COLOR)
    'Alterar a cor inicial de fundo do botão
    UserControl.BackColor = New_Value
    PropertyChanged "BackGround"
End Property

Public Property Get BackColorNormal() As OLE_COLOR
    'Escolher a cor inicial de fundo do botão
    BackColorNormal = Shape1.FillColor
End Property

Public Property Let BackColorNormal(New_Value As OLE_COLOR)
    'Alterar a cor inicial de fundo do botão
    cor_fundo_normal = New_Value
    Shape1.FillColor = New_Value
    PropertyChanged "BackColorNormal"
End Property

Public Property Get BackColorHover() As OLE_COLOR
    'Escolher a cor inicial de fundo do botão
    BackColorHover = cor_fundo_hover
End Property

Public Property Let BackColorHover(New_Value As OLE_COLOR)
    'Alterar a cor inicial de fundo do botão
    cor_fundo_hover = New_Value
    'Shape1.FillColor = new_Value
    PropertyChanged "BackColorHover"
End Property

Public Property Get BackColorDown() As OLE_COLOR
    'Escolher a cor inicial de fundo do botão
    BackColorDown = cor_fundo_down
End Property

Public Property Let BackColorDown(New_Value As OLE_COLOR)
    'Alterar a cor inicial de fundo do botão
    cor_fundo_down = New_Value
    'Shape1.FillColor = new_Value
    PropertyChanged "BackColorDown"
End Property

Public Property Get BorderCustom() As OLE_COLOR
    'Escolher a nova cor do contorno do botao
    BorderCustom = cor_contorno_custom
End Property

Public Property Let BorderCustom(New_Value As OLE_COLOR)
    'Alterar o contorno do botao
    'Shape1.BorderCustom = new_Value
    cor_contorno_custom = New_Value
    PropertyChanged "BorderCustom"
End Property

Public Property Get BorderColorNormal() As OLE_COLOR
    'Escolher a cor inicial de fundo do botão
    BorderColorNormal = Shape1.BorderColor
End Property

Public Property Let BorderColorNormal(New_Value As OLE_COLOR)
    'Alterar a cor inicial de fundo do botão
    cor_contorno_normal = New_Value
    Shape1.BorderColor = New_Value
    PropertyChanged "BorderColorNormal"
End Property

Public Property Get BorderColorHover() As OLE_COLOR
    'Escolher a cor inicial da letra do botão
    BorderColorHover = cor_contorno_hover
End Property

Public Property Let BorderColorHover(New_Value As OLE_COLOR)
    'Alterar a cor inicial da letra do botão
    cor_contorno_hover = New_Value
    'Shape1.FillColor = new_Value
    PropertyChanged "BorderColorHover"
End Property

Public Property Get BorderColorDown() As OLE_COLOR
    'Escolher a cor inicial da letra do bot
    BorderColorDown = cor_contorno_down
End Property

Public Property Let BorderColorDown(New_Value As OLE_COLOR)
    'Alterar a cor inicial da letra do bot
    cor_contorno_down = New_Value
    'Shape1.FillColor = new_Value
    PropertyChanged "BorderColorDown"
End Property
'
'Public Property Get AlignementHorizontal() As AlignmentConstants
'    AlignementHorizontal = txt.Alignment
'End Property
'
'Public Property Let AlignementHorizontal(ByVal newValue As AlignmentConstants)
'    If newValue > 2 Then newValue = 2
'    txt.Alignment = newValue
'    lbl.Alignment = txt.Alignment
'    PropertyChanged ("AlignementHorizontal")
'End Property


