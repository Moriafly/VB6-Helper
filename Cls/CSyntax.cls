VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSyntax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
'*��ģ�����ƣ�CSyntax
'*��ģ�鹦�ܣ��﷨����
'*˵����������˸Ķ������ʸ���һ��(progame@cnnb.net)
'*      ��ҳ��http://www.NewXing.com
'*���ߣ�progame  2002-10-07  17:54:44
'***************************************************************
Option Explicit
Option Compare Binary

Public Event Progress(Value As Integer) '*��������

Private m_sTitle            As String   '*�﷨�ļ�Title

Private m_sLnCommFlag1      As String   '*��ע��ͷһ
Private m_sLnCommFlag2      As String   '*��ע��ͷ��
Private m_sBlkCommHead1     As String   '*��ע��ͷһ
Private m_sBlkCommFoot1     As String   '*��ע��βһ
Private m_sBlkCommHead2     As String   '*��ע��ͷ��
Private m_sBlkCommFoot2     As String   '*��ע��β��
Private m_objCommFont       As CFont    '*ע����������

Private m_sCommRtfHead      As String   '*ע������RTF��ʽͷ
Private m_sCommRtfFoot      As String   '*ע������RTF��ʽβ

Private m_objNumFont        As CFont    '*������������

Private m_sNumRtfHead       As String   '*��������RTF��ʽͷ
Private m_sNumRtfFoot       As String   '*��������RTF��ʽβ

Private m_iKwdCount         As Integer  '*�ؼ���������
Private m_sKwdSet()         As String   '*����һ��ؼ��ֵ�ȫ�����(��chr(3)�ָ�)
Private m_sKwdTitle()       As String   '*�˹ؼ������͵�����
Private m_objKwdFont()      As CFont    '*�˹ؼ������͵���������
Private m_sKwdRtfHead()     As String   '*�ؼ�������RTF��ʽͷ
Private m_sKwdRtfFoot()     As String   '*�ؼ�������RTF��ʽβ

Private m_bCase             As Boolean  '*�Ƿ��С������

Private m_sQuotFlag1        As String   '*�ַ����߽�һ
Private m_sQuotFlag2        As String   '*�ַ����߽��
Private m_bMultiLine        As Boolean  '*�Ƿ�������Ч
Private m_objQuotFont       As CFont    '*�ַ�����������
Private m_sQuotRtfHead      As String   '*�ַ�������RTF��ʽͷ
Private m_sQuotRtfFoot      As String   '*�ַ�������RTF��ʽβ

Private m_sEscape           As String   '*����������ַ������ַ������ַ�����

Private m_objColorManager   As CColor   '*��ɫ������
Private m_sColorRtf         As String   '*RTF����ɫ��Ϣ

Private m_sWord()         As String     '*��һ���ַ����ָ���һ��������

Private m_iQuotState        As Integer  '*��ǰ�Ƿ����ַ�����[0..�����ַ���/1..���ַ���һ��/2..���ַ�������]
Private m_iCommState        As Integer  '*��ǰע��״̬[0..��ע��/1..����ע��һ/2..����ע�Ͷ�/3..����ע��һ/4..����ע�Ͷ�]


Private Const DELIMITER = ",(){}[]-+*%/='~!&|\<>?:;."


Public Function ReadFile(filename As String) As Boolean
'*���﷨�ļ�����
Dim fn      As Integer
Dim s       As String
Dim i       As Integer
    fn = FreeFile()
    
    Open filename For Input As #fn
        Do While Not EOF(fn)
            Line Input #fn, s
            Call PreFormat(s)
            Call ReadDetail(fn, s)
        Loop
    Close #fn
    
    '*��������ַ���Ԥ������
    m_sQuotRtfHead = m_objQuotFont.getRtfHead(m_objColorManager)
    m_sQuotRtfFoot = m_objQuotFont.getRtfTail
    m_sCommRtfHead = m_objCommFont.getRtfHead(m_objColorManager)
    m_sCommRtfFoot = m_objCommFont.getRtfTail
    
    m_sNumRtfHead = m_objNumFont.getRtfHead(m_objColorManager)
    m_sNumRtfFoot = m_objNumFont.getRtfTail
    
    m_sColorRtf = m_objColorManager.ComStr
    
    For i = 1 To m_iKwdCount
        m_sKwdRtfHead(i) = m_objKwdFont(i).getRtfHead(m_objColorManager)
        m_sKwdRtfFoot(i) = m_objKwdFont(i).getRtfTail
    Next i
    
    ReadFile = True

    Exit Function
    
Err_Proc:
    MsgBox Err.Description
    ReadFile = False
End Function

Private Sub PreFormat(ByRef s As String)
'*���﷨�ļ�Ԥ��������Ϊ��ʱ�������Ƚ��и�ʽ������Ӧ"}" "{" "\"
    s = Replace(Replace(Replace(s, "\", "\\"), "{", "\{"), "}", "\}")
End Sub

Private Sub ReadDetail(filenum As Integer, s As String)
'*��ȡ�﷨�ļ��ľ���ʵ��
Dim sTmp    As String
Dim i       As Integer
Dim lstr    As String
Dim rstr    As String

    s = Trim(s)
    If Len(s) = 0 Then              'blank
        Exit Sub
    End If
    If Left(s, 1) = ";" Then        'comment
        Exit Sub
    End If
    
    '*��ȡ�ַ�������������
    If Len(s) = 10 And Left(s, 10) = "#QUOTATION" Then
        Do While Not EOF(filenum)
            Line Input #filenum, sTmp
            Call PreFormat(sTmp)
            If Trim(sTmp) = "" Or Trim(sTmp) = "#" Or Left(sTmp, 1) = ";" Then
                Exit Sub
            End If
            '��ȡ��ϸ����
            i = InStr(sTmp, "=")
            If Not (i = 0 Or (i = 1 And Len(sTmp) = i)) Then
                Call ReadFont(sTmp, i, m_objQuotFont)
            End If
        Loop
        
        Exit Sub
    End If
    
     '*��ȡע�͵���������
    If Len(s) = 8 And UCase(Left(s, 8)) = "#COMMENT" Then
        Do While Not EOF(filenum)
            Line Input #filenum, sTmp
            Call PreFormat(sTmp)
            If Trim(sTmp) = "" Or Trim(sTmp) = "#" Or Left(sTmp, 1) = ";" Then
                Exit Sub
            End If
            '��ȡ��ϸ����
            i = InStr(sTmp, "=")
            If Not (i = 0 Or (i = 1 And Len(sTmp) = i)) Then
                Call ReadFont(sTmp, i, m_objCommFont)
            End If
        Loop
        
        Exit Sub
    End If
    
     '*��ȡ���ֵ���������
    If Len(s) = 7 And UCase(Left(s, 7)) = "#NUMBER" Then
        Do While Not EOF(filenum)
            Line Input #filenum, sTmp
            Call PreFormat(sTmp)
            If Trim(sTmp) = "" Or Trim(sTmp) = "#" Or Left(sTmp, 1) = ";" Then
                Exit Sub
            End If
            '��ȡ��ϸ����
            i = InStr(sTmp, "=")
            If Not (i = 0 Or (i = 1 And Len(sTmp) = i)) Then
                Call ReadFont(sTmp, i, m_objNumFont)
            End If
        Loop
        
        Exit Sub
    End If
    
    '*��ȡ�ؼ���
    If UCase(Left(s, 8)) = "#KEYWORD" Then  'keyword head
    
        sTmp = Right(s, Len(s) - 9)
        m_iKwdCount = m_iKwdCount + 1
        
        ReDim Preserve m_sKwdSet(1 To m_iKwdCount)
        ReDim Preserve m_sKwdTitle(1 To m_iKwdCount)
        ReDim Preserve m_objKwdFont(1 To m_iKwdCount)
        ReDim Preserve m_sKwdRtfHead(1 To m_iKwdCount)
        ReDim Preserve m_sKwdRtfFoot(1 To m_iKwdCount)
        
        Set m_objKwdFont(m_iKwdCount) = New CFont
        
        m_sKwdTitle(m_iKwdCount) = sTmp
        m_sKwdSet(m_iKwdCount) = Chr(3)
        
        Do While Not EOF(filenum)
            Line Input #filenum, sTmp
            Call PreFormat(sTmp)
            If Trim(sTmp) = "" Or Trim(sTmp) = "#" Or Left(sTmp, 1) = ";" Then
                Exit Sub
            End If

            '��ȡ�ؼ��ֵ���ϸ����
            i = InStr(sTmp, "=")
            If i = 0 Or (i = 1 And Len(sTmp) = i) Then
                m_sKwdSet(m_iKwdCount) = m_sKwdSet(m_iKwdCount) & IIf(m_bCase, Trim(sTmp), UCase(Trim(sTmp))) & Chr(3)
            Else
                Call ReadFont(sTmp, i, m_objKwdFont(m_iKwdCount))
            End If
        Loop
        
        Exit Sub
    End If
    
    i = InStr(s, "=")
    If i = 0 Or i = 1 Or Len(s) = i Then
        Exit Sub
    End If
    
    '*��ȡ��������
    lstr = Trim(Left(s, i - 1))
    rstr = Trim(Right(s, Len(s) - i))
    If lstr = "" Or rstr = "" Then
        Exit Sub
    End If
    Select Case UCase(lstr)
        Case "#TITLE"
            rstr = rstr
        Case "#QUOTATION1"
            m_sQuotFlag1 = rstr
        Case "#QUOTATION2"
            m_sQuotFlag2 = rstr
        Case "#CONTINUE_QUOTE"
            m_bMultiLine = IIf(LCase(rstr) = "y", True, False)
        Case "#LINECOMMENT"
            m_sLnCommFlag1 = rstr
        Case "#LINECOMMENT2"
            m_sLnCommFlag2 = rstr
        Case "#COMMENTON"
            m_sBlkCommHead1 = rstr
        Case "#COMMENTOFF"
            m_sBlkCommFoot1 = rstr
        Case "#COMMENTON2"
            m_sBlkCommHead2 = rstr
        Case "#COMMENTOFF2"
            m_sBlkCommFoot2 = rstr
        Case "#CASE"
            m_bCase = IIf(LCase(rstr) = "y", True, False)
            
        Case "#ESCAPE"
            m_sEscape = rstr
        Case Else
           
    End Select
End Sub

Private Sub ReadFont(sValue As String, i As Integer, Font As CFont)
'*��ȡ�ļ��е���������
'*      i       --"="����λ��
'*      font    --Ҫ�����������

Dim lstr        As String
Dim rstr        As String
    lstr = Trim(Left(sValue, i - 1))
    rstr = Trim(Right(sValue, Len(sValue) - i))
    If lstr <> "" And rstr <> "" Then
        Select Case LCase(lstr)
            Case "forecolor"
                Font.bForecolor = True
                Font.ForeColor = rstr
                m_objColorManager.AddColor CLng(rstr)
            Case "backcolor"
                Font.bBackcolor = True
                Font.BackColor = rstr
                m_objColorManager.AddColor CLng(rstr)
            Case "bold"
                Font.bold = IIf(LCase(rstr) = "y", True, False)
            Case "italic"
                Font.italic = IIf(LCase(rstr) = "y", True, False)
            Case "underline"
                Font.underline = IIf(LCase(rstr) = "y", True, False)
            Case "strike"
                Font.strike = IIf(LCase(rstr) = "y", True, False)
            Case Else
                
        End Select
    End If
End Sub

Private Sub Class_Initialize()
    Set m_objColorManager = New CColor
    Set m_objCommFont = New CFont
    Set m_objQuotFont = New CFont
    Set m_objNumFont = New CFont
End Sub

Private Sub Class_Terminate()
Dim i   As Integer
    Set m_objColorManager = Nothing
    Set m_objCommFont = Nothing
    Set m_objQuotFont = Nothing
    Set m_objNumFont = Nothing
    
    For i = 1 To m_iKwdCount
        Set m_objKwdFont(i) = Nothing
    Next i
    
    Erase m_objKwdFont
    Erase m_sKwdTitle
    Erase m_sKwdSet
    Erase m_sKwdRtfHead
    Erase m_sKwdRtfFoot
    Erase m_sWord
End Sub

'**************************************************************
'*HighLightRichEdit
'*���ܣ��Կؼ������﷨������ʾ
'*˵����
'*���������
'*    ctrl      ���ͣ�RichTextBox
'*���ز�����
'*
'*���ߣ�progame  ���ڣ�2002-09-29  17:38:22
'**************************************************************
Public Sub HighLightRichEdit(ctrl As RichTextBox)
Dim sFileName   As String

    sFileName = TempFileName
    HighLightCtrl ctrl, sFileName, True
    
End Sub


'**************************************************************
'*HighLightCtrl
'*���ܣ��Կؼ������﷨������ʾ
'*˵����
'*���������
'*    ctrl           ���ͣ�RichTextBox      --�ؼ�
'*    filename       ���ͣ�String           --��ʱ�ļ���
'*    bKillFile      ���ͣ�Boolean          --��ɺ��Ƿ�ɾ���ļ�
'*���ز�����
'*
'*���ߣ�progame  ���ڣ�2002-09-29  17:38:22
'**************************************************************
Private Sub HighLightCtrl(ctrl As RichTextBox, filename As String, bKillFile As Boolean)
Dim sLine() As String
Dim i       As Long
Dim l       As Long
Dim sTmp        As String
Dim iProgress   As Integer      '*��ǰ��������
Dim iTmp        As Integer

    
    sTmp = ctrl.Text
    l = Len(sTmp)
    If l = 0 Then
        Exit Sub
    End If
    '*�Ƚ�
    Open filename For Output As #1
    
        Print #1, HEAD_FIRST & m_sColorRtf & HEAD_LAST
 
        m_iQuotState = 0
        m_iCommState = 0
        
        iProgress = 0
        
        '*���ڴ��ı����зֿ鴦����ÿ1KΪһ����λ
        Dim aTmp()      As String
        Dim j           As Long
        Dim lstPos      As Long
        lstPos = 1
        ReDim aTmp(1 To l / 4000 + 1)
        For i = 1 To UBound(aTmp)
            j = InStr(lstPos + 4000, sTmp, vbCrLf, vbTextCompare)
            If j = 0 Then
                j = l
                aTmp(i) = Mid(sTmp, lstPos, j - lstPos + 1)
                lstPos = -1
            Else
                aTmp(i) = Mid(sTmp, lstPos, j - lstPos + 1)
                lstPos = j + 1
            End If
            '*���������ַ��ͷָ��ַ�
            aTmp(i) = Replace(aTmp(i), "\", "\\")
            aTmp(i) = Replace(aTmp(i), "{", "\{")
            aTmp(i) = Replace(aTmp(i), "}", "\}")
            aTmp(i) = Replace(aTmp(i), "(", " ( ")
            aTmp(i) = Replace(aTmp(i), ")", " ) ")
            aTmp(i) = Replace(aTmp(i), ";", " ; ")
            aTmp(i) = Replace(aTmp(i), "#", " # ")
            aTmp(i) = Replace(aTmp(i), Chr(9), " " & Chr(9) & " ")
            aTmp(i) = Replace(aTmp(i), vbTab, " " & vbTab & " ")
        
            sLine = Split(aTmp(i), vbCrLf)
            
                
            For j = 0 To UBound(sLine)
            
                Call HighLightLine(sLine(j))    '*����һ��
                
                '*�ָ��ָ��ַ�
                aTmp(i) = Replace(sLine(j), " ( ", "(")
                aTmp(i) = Replace(aTmp(i), " ) ", ")")
                aTmp(i) = Replace(aTmp(i), " ; ", ";")
                aTmp(i) = Replace(aTmp(i), " # ", "#")
                aTmp(i) = Replace(aTmp(i), " " & Chr(9) & " ", Chr(9))
                aTmp(i) = Replace(aTmp(i), " " & vbTab & " ", vbTab)
                
                Print #1, aTmp(i) & "\par "
                
                
    
            Next j
            
            If lstPos = -1 Then
                Exit For
            End If
            
            iTmp = i * 100 / UBound(aTmp)
            If iProgress <> iTmp Then
                iProgress = iTmp
                RaiseEvent Progress(iProgress)
            End If
                
        Next i
        
        Print #1, "}"

    Close #1

    ctrl.LoadFile filename, rtfRTF
    
    Erase sLine
    
    If Dir(filename) <> "" And bKillFile Then
        Kill filename
    End If
End Sub

Private Sub HighLightLine(ByRef sLine As String)
'*ɨ���ַ��������ʽ������ַ���
Dim i       As Long
Dim j       As Integer
Dim k       As Integer
Dim bComm As Boolean

    m_sWord = Split(sLine, " ")
    sLine = ""
    
    k = UBound(m_sWord)
    
    For i = 0 To k

        '*��������ַ�����������Ƚ���ע���жϣ��ٽ��йؼ����ж�
        If m_iQuotState = 0 Then
        
            bComm = TreatComment(m_sWord(i), (i = k))         '*����ע��
        
        End If

        If Not bComm Then

                If TreatQuot(m_sWord(i), (i = k), 1) Then       '*�����ַ���
                
                    '*�ָ������������ַ�����־
                    If m_sQuotFlag1 <> "" Then
                        m_sWord(i) = Replace(m_sWord(i), Chr(1), m_sQuotFlag1)
                    End If
                    If m_sQuotFlag2 <> "" Then
                        m_sWord(i) = Replace(m_sWord(i), Chr(2), m_sQuotFlag2)
                    End If
                Else
                
                    If m_iQuotState = 0 Then
                        Call TreatKeyWord(m_sWord(i))         '*�����ؼ���
                    End If
                    
                End If

        End If
        
        If i = k Then
            sLine = sLine & m_sWord(i)
        Else
            sLine = sLine & m_sWord(i) & " "
        End If
        
    Next i

End Sub

Private Function TreatComment(ByRef sWord As String, bLast As Boolean) As Boolean
'*����ע��
'*      bLast       --�Ƿ����һ���ַ�
'*      ����ֵ��    sWord�Ƿ�ע��
    Select Case m_iCommState
        Case 0
            If m_sLnCommFlag1 <> "" Then
                If InStr(1, sWord, m_sLnCommFlag1, vbBinaryCompare) = 1 Then  '*�ǵ���ע��һ
                    m_iCommState = 1
                    '*����ע������RTF��ʽͷ
                    sWord = m_sCommRtfHead & sWord
                    TreatComment = True
                    '*������һ�У��ٽ��д���
                    If bLast Then
                        TreatComment sWord, True
                    End If
                    Exit Function
                End If
            End If
            If m_sLnCommFlag2 <> "" Then
                If InStr(1, sWord, m_sLnCommFlag2, vbBinaryCompare) = 1 Then  '*�ǵ���ע�Ͷ�
                    m_iCommState = 2
                    '*����ע������RTF��ʽͷ
                    sWord = m_sCommRtfHead & sWord
                    TreatComment = True
                    '*������һ���ʣ��ٽ��д���
                    If bLast Then
                        TreatComment sWord, True
                    End If
                    Exit Function
                End If
            End If
            If m_sBlkCommHead1 <> "" And m_sBlkCommFoot1 <> "" Then
                If InStr(1, sWord, m_sBlkCommHead1, vbBinaryCompare) = 1 Then  '*�Ƕ���ע��һ
                    m_iCommState = 3
                    '*����ע������RTF��ʽͷ
                    sWord = m_sCommRtfHead & sWord
                    TreatComment = True
                    TreatComment sWord, bLast
                    Exit Function
                End If
            End If
            If m_sBlkCommHead2 <> "" And m_sBlkCommFoot2 <> "" Then
                If InStr(1, sWord, m_sBlkCommHead2, vbBinaryCompare) = 1 Then '*�Ƕ���ע�Ͷ�
                    m_iCommState = 4
                    '*����ע������RTF��ʽͷ
                    sWord = m_sCommRtfHead & sWord
                    TreatComment = True
                    TreatComment sWord, bLast
                    Exit Function
                End If
            End If
        Case 1, 2
            TreatComment = True
            If bLast Then       '*����ע�ͣ��ҵ�ǰΪ�������һ���ʣ���ע�ͱ�־�ָ�Ϊ0
                sWord = sWord & m_sCommRtfFoot
                m_iCommState = 0
            End If
        Case 3
            TreatComment = True
            If m_sBlkCommHead1 <> "" And m_sBlkCommFoot1 <> "" And Len(sWord) > 1 Then
                If InStrRev(sWord, m_sBlkCommFoot1, , vbBinaryCompare) <> 0 Then '*�Ƕ���ע��һ
                    m_iCommState = 0
                    '*����ע������RTF��ʽβ
                    sWord = sWord & m_sCommRtfFoot
                    Exit Function
                End If
            End If
        Case 4
            TreatComment = True
            If m_sBlkCommHead2 <> "" And m_sBlkCommFoot2 <> "" And Len(sWord) > 1 Then
                If InStrRev(sWord, m_sBlkCommFoot2, , vbBinaryCompare) <> 0 Then '*�Ƕ���ע�Ͷ�
                    m_iCommState = 0
                    '*����ע������RTF��ʽβ
                    sWord = sWord & m_sCommRtfFoot
                    Exit Function
                End If
            End If
        Case Else
        
    End Select
    
End Function

Private Function TreatKeyWord(ByRef sWord As String) As Boolean
'*�����ؼ���
'*      ����ֵ��    sWord�Ƿ�ؼ���
Dim i           As Integer
Dim bKeyWord    As Boolean
    '* �ȴ����ǲ�������
    If IsNumeric(sWord) Then
        sWord = m_sNumRtfHead & sWord & m_sNumRtfFoot
        Exit Function
    End If
    
    For i = 1 To m_iKwdCount
        If m_bCase Then
            If (InStr(1, m_sKwdSet(i), Chr(3) & sWord & Chr(3), vbBinaryCompare) <> 0) Then
               bKeyWord = True
            End If
        Else
            If (InStr(1, m_sKwdSet(i), Chr(3) & UCase(sWord) & Chr(3), vbBinaryCompare) <> 0) Then
                bKeyWord = True
            End If
        End If
        
        If bKeyWord Then
            '*��������RTF��ʽ��ͷβ�ڴ�����
            sWord = m_sKwdRtfHead(i) & sWord & m_sKwdRtfFoot(i)
            TreatKeyWord = True
            Exit Function
        End If
    Next i
End Function

Private Function TreatQuot(ByRef sWord As String, bLast As Boolean, iBegin As Integer) As Boolean
'*�����ַ���
'*      bLast       �Ƿ�������һ����
'*      iBegin      ����λ��
'*      ����ֵ��    �Ƿ������ַ�������
Dim i       As Integer
Dim j       As Integer
Dim sQuot   As String       '*��ǰ�����жϵ��ַ����߽��
    Select Case m_iQuotState
        Case 0
            '*�ж������ַ�����־�ȳ��֣�ʹ���ȳ��ֵ��ַ�����־�������ַ�����ʼλ��
            i = 0
            j = 0
            If m_sQuotFlag1 <> "" Then
                i = InStr(iBegin, sWord, m_sQuotFlag1, vbBinaryCompare)
            End If
            If m_sQuotFlag2 <> "" Then
                j = InStr(iBegin, sWord, m_sQuotFlag2, vbBinaryCompare)
            End If
            If i = 0 And j = 0 Then     '*û���ַ�����־���֣������ж�
                Exit Function
            End If
            If (j = 0) Or (i <> 0 And i < j) Then
                sQuot = m_sQuotFlag1
            End If
            If (j <> 0 And j < i) Or (i = 0) Then
                sQuot = m_sQuotFlag2
            End If
        Case 1
            sQuot = m_sQuotFlag1
        Case 2
            sQuot = m_sQuotFlag2
        Case Else
    End Select
    
    '*���ҵ�ǰ���ַ�����־
    i = InStr(iBegin, sWord, sQuot, vbBinaryCompare)
    
    '*���ǰһ���ַ��������ַ�ǰ׺����˴β�����Ч
    If m_sEscape <> "" Then
        If i > 1 Then
            j = InStr(iBegin, sWord, m_sEscape)
            If j <> 0 And j = i - Len(m_sEscape) Then
                '*�����ַ�����־����
                If sQuot = m_sQuotFlag1 Then
                    sWord = Replace(sWord, sQuot, Chr(1), , 1)
                Else
                    sWord = Replace(sWord, sQuot, Chr(2), , 1)
                End If
                TreatQuot = True
                '*�����ݹ鴦��
                Call TreatQuot(sWord, bLast, i + 1)
                Exit Function
            End If
        End If
    End If
    
    If i <> 0 Then
    
        TreatQuot = True
        
        '*��ʽ���ַ���
        If sQuot = m_sQuotFlag1 Then
            If m_iQuotState = 0 Then
                sWord = Replace(sWord, sQuot, m_sQuotRtfHead & Chr(1), , 1)
                m_iQuotState = 1
            Else
                sWord = Replace(sWord, sQuot, Chr(1) & m_sQuotRtfFoot, , 1)
                m_iQuotState = 0
            End If
        Else
            If m_iQuotState = 0 Then
                sWord = Replace(sWord, sQuot, m_sQuotRtfHead & Chr(2), , 1)
                m_iQuotState = 2
            Else
                sWord = Replace(sWord, sQuot, Chr(2) & m_sQuotRtfFoot, , 1)
                m_iQuotState = 0
            End If
        End If

        '*�����ݹ鴦��
        Call TreatQuot(sWord, bLast, i)
        
    Else
    
        '*����Ѿ�û���˱�־����������Ч������£���ǰ�����Ϊ���һ��ʱ��ǿ�н���
        If bLast And Not m_bMultiLine And m_iQuotState <> 0 Then
            sWord = sWord & m_sQuotRtfFoot
            m_iQuotState = 0
        End If
        
    End If
    

End Function


'**************************************************************
'*Rtf2Html
'*���ܣ�    ��RTF�ؼ�����ת��ΪHtml�ַ���
'*˵����
'*���������
'*    rTxt      ���ͣ�RichTextBox
'*���ز�����
'*      HTML�ַ���
'*���ߣ�progame  ���ڣ�2002-10-17  17:36:17
'**************************************************************
Public Function Rtf2Html(ctrl As RichTextBox) As String
Dim byt()       As Byte
Dim sFileName   As String
    sFileName = TempFileName
    
    HighLightCtrl ctrl, sFileName, False
    
    '*�����ļ�
    Open sFileName For Binary As #1
        ReDim byt(1 To FileLen(sFileName))
        Get #1, , byt
    Close #1
    
    If Dir(sFileName) <> "" Then
        Kill sFileName
    End If
    
    Rtf2Html = StrConv(byt, vbUnicode)
    
    '*�������ַ��ָ�
    Rtf2Html = Replace(Rtf2Html, "\\", "\")
    Rtf2Html = Replace(Rtf2Html, "\{", "{")
    Rtf2Html = Replace(Rtf2Html, "\}", "}")
    
    '*��ɾ��RTF�ļ�ͷ
    Rtf2Html = Replace(Rtf2Html, HEAD_FIRST & m_sColorRtf & HEAD_LAST, "", , 1)
    '*ת�����з���
    Rtf2Html = Replace(Rtf2Html, "\par ", "")
    
    '*�������ַ�����
    Rtf2Html = Replace(Rtf2Html, "&", "&amp;")
    Rtf2Html = Replace(Rtf2Html, "<", "&lt;")
    Rtf2Html = Replace(Rtf2Html, ">", "&gt;")
    '*ת��ע���������
    FontRtf2Html Rtf2Html, m_objQuotFont
    FontRtf2Html Rtf2Html, m_objCommFont
    Dim i       As Integer
    For i = 1 To m_iKwdCount
        FontRtf2Html Rtf2Html, m_objKwdFont(i)
    Next i
    
    '*ȥ�����һ��"}"
    Rtf2Html = Left(Rtf2Html, Len(Rtf2Html) - 4)
    
    
    '*����HTMLͷβ
    Rtf2Html = HEAD_HTML & Rtf2Html & TAIL_HTML
 
End Function

Private Sub FontRtf2Html(ByRef sDeal As String, ByRef objFont As CFont)
'*��sDeal�ַ�������RTF��HTML������ת��
    sDeal = Replace(sDeal, objFont.getRtfHead(m_objColorManager), objFont.getHtmlHead(m_objColorManager))
    sDeal = Replace(sDeal, objFont.getRtfTail, objFont.getHtmlTail)
End Sub

'**************************************************************
'*Rtf2Ubb
'*���ܣ�    ��RTF�ؼ�����ת��ΪUbb�ַ���
'*˵����
'*���������
'*    rTxt      ���ͣ�RichTextBox
'*���ز�����
'*      HTML�ַ���
'*���ߣ�progame  ���ڣ�2002-10-17  17:36:17
'**************************************************************
Public Function Rtf2Ubb(ctrl As RichTextBox) As String
Dim byt()       As Byte
Dim sFileName   As String
    sFileName = TempFileName
    
    HighLightCtrl ctrl, sFileName, False

    '*�����ļ�
    Open sFileName For Binary As #1
        ReDim byt(1 To FileLen(sFileName))
        Get #1, , byt
    Close #1
    
    If Dir(sFileName) <> "" Then
        Kill sFileName
    End If
    
    Rtf2Ubb = StrConv(byt, vbUnicode)
    
    '*�������ַ��ָ�
    Rtf2Ubb = Replace(Rtf2Ubb, "\\", "\")
    Rtf2Ubb = Replace(Rtf2Ubb, "\{", "{")
    Rtf2Ubb = Replace(Rtf2Ubb, "\}", "}")
    
    '*��ɾ��RTF�ļ�ͷ
    Rtf2Ubb = Replace(Rtf2Ubb, HEAD_FIRST & m_sColorRtf & HEAD_LAST, "", , 1)
    '*ת�����з���
    Rtf2Ubb = Replace(Rtf2Ubb, "\par ", vbCrLf)
    
    '*ת��ע���������
    FontRtf2Ubb Rtf2Ubb, m_objQuotFont
    FontRtf2Ubb Rtf2Ubb, m_objCommFont
    Dim i       As Integer
    For i = 1 To m_iKwdCount
        FontRtf2Ubb Rtf2Ubb, m_objKwdFont(i)
    Next i
    
    '*ȥ�����һ��"}"
    Rtf2Ubb = Left(Rtf2Ubb, Len(Rtf2Ubb) - 4)
    
 
End Function

Private Sub FontRtf2Ubb(ByRef sDeal As String, ByRef objFont As CFont)
'*��sDeal�ַ�������RTF��UBB������ת��
    sDeal = Replace(sDeal, objFont.getRtfHead(m_objColorManager), objFont.getUBBHead(m_objColorManager))
    sDeal = Replace(sDeal, objFont.getRtfTail, objFont.getUBBTail)
End Sub