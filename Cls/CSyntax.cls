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
'*类模块名称：CSyntax
'*类模块功能：语法高亮
'*说明：如果做了改动，请邮给我一份(progame@cnnb.net)
'*      主页：http://www.NewXing.com
'*作者：progame  2002-10-07  17:54:44
'***************************************************************
Option Explicit
Option Compare Binary

Public Event Progress(Value As Integer) '*处理进度

Private m_sTitle            As String   '*语法文件Title

Private m_sLnCommFlag1      As String   '*行注释头一
Private m_sLnCommFlag2      As String   '*行注释头二
Private m_sBlkCommHead1     As String   '*块注释头一
Private m_sBlkCommFoot1     As String   '*块注释尾一
Private m_sBlkCommHead2     As String   '*块注释头二
Private m_sBlkCommFoot2     As String   '*块注释尾二
Private m_objCommFont       As CFont    '*注释字体设置

Private m_sCommRtfHead      As String   '*注释字体RTF格式头
Private m_sCommRtfFoot      As String   '*注释字体RTF格式尾

Private m_objNumFont        As CFont    '*数字字体设置

Private m_sNumRtfHead       As String   '*数字字体RTF格式头
Private m_sNumRtfFoot       As String   '*数字字体RTF格式尾

Private m_iKwdCount         As Integer  '*关键字类型数
Private m_sKwdSet()         As String   '*存入一类关键字的全部组合(以chr(3)分隔)
Private m_sKwdTitle()       As String   '*此关键字类型的名称
Private m_objKwdFont()      As CFont    '*此关键字类型的字体设置
Private m_sKwdRtfHead()     As String   '*关键字字体RTF格式头
Private m_sKwdRtfFoot()     As String   '*关键字字体RTF格式尾

Private m_bCase             As Boolean  '*是否大小字敏感

Private m_sQuotFlag1        As String   '*字符串边界一
Private m_sQuotFlag2        As String   '*字符串边界二
Private m_bMultiLine        As Boolean  '*是否续行有效
Private m_objQuotFont       As CFont    '*字符串字体设置
Private m_sQuotRtfHead      As String   '*字符串字体RTF格式头
Private m_sQuotRtfFoot      As String   '*字符串字体RTF格式尾

Private m_sEscape           As String   '*后面跟敏感字符，此字符后面字符忽略

Private m_objColorManager   As CColor   '*颜色管理类
Private m_sColorRtf         As String   '*RTF内颜色信息

Private m_sWord()         As String     '*对一行字符串分隔成一个词数组

Private m_iQuotState        As Integer  '*当前是否在字符串内[0..不是字符串/1..在字符串一内/2..在字符串二内]
Private m_iCommState        As Integer  '*当前注释状态[0..无注释/1..单行注释一/2..单行注释二/3..多行注释一/4..多行注释二]


Private Const DELIMITER = ",(){}[]-+*%/='~!&|\<>?:;."


Public Function ReadFile(filename As String) As Boolean
'*将语法文件读入
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
    
    '*对字体的字符串预先生成
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
'*进语法文件预处理，因为到时候代码会先进行格式化以适应"}" "{" "\"
    s = Replace(Replace(Replace(s, "\", "\\"), "{", "\{"), "}", "\}")
End Sub

Private Sub ReadDetail(filenum As Integer, s As String)
'*读取语法文件的具体实现
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
    
    '*读取字符串的字体配置
    If Len(s) = 10 And Left(s, 10) = "#QUOTATION" Then
        Do While Not EOF(filenum)
            Line Input #filenum, sTmp
            Call PreFormat(sTmp)
            If Trim(sTmp) = "" Or Trim(sTmp) = "#" Or Left(sTmp, 1) = ";" Then
                Exit Sub
            End If
            '读取详细内容
            i = InStr(sTmp, "=")
            If Not (i = 0 Or (i = 1 And Len(sTmp) = i)) Then
                Call ReadFont(sTmp, i, m_objQuotFont)
            End If
        Loop
        
        Exit Sub
    End If
    
     '*读取注释的字体配置
    If Len(s) = 8 And UCase(Left(s, 8)) = "#COMMENT" Then
        Do While Not EOF(filenum)
            Line Input #filenum, sTmp
            Call PreFormat(sTmp)
            If Trim(sTmp) = "" Or Trim(sTmp) = "#" Or Left(sTmp, 1) = ";" Then
                Exit Sub
            End If
            '读取详细内容
            i = InStr(sTmp, "=")
            If Not (i = 0 Or (i = 1 And Len(sTmp) = i)) Then
                Call ReadFont(sTmp, i, m_objCommFont)
            End If
        Loop
        
        Exit Sub
    End If
    
     '*读取数字的字体配置
    If Len(s) = 7 And UCase(Left(s, 7)) = "#NUMBER" Then
        Do While Not EOF(filenum)
            Line Input #filenum, sTmp
            Call PreFormat(sTmp)
            If Trim(sTmp) = "" Or Trim(sTmp) = "#" Or Left(sTmp, 1) = ";" Then
                Exit Sub
            End If
            '读取详细内容
            i = InStr(sTmp, "=")
            If Not (i = 0 Or (i = 1 And Len(sTmp) = i)) Then
                Call ReadFont(sTmp, i, m_objNumFont)
            End If
        Loop
        
        Exit Sub
    End If
    
    '*读取关键字
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

            '读取关键字的详细内容
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
    
    '*读取基本配置
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
'*读取文件中的字体设置
'*      i       --"="所在位置
'*      font    --要传入的字体类

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
'*功能：对控件进行语法高亮显示
'*说明：
'*传入参数：
'*    ctrl      类型：RichTextBox
'*返回参数：
'*
'*作者：progame  日期：2002-09-29  17:38:22
'**************************************************************
Public Sub HighLightRichEdit(ctrl As RichTextBox)
Dim sFileName   As String

    sFileName = TempFileName
    HighLightCtrl ctrl, sFileName, True
    
End Sub


'**************************************************************
'*HighLightCtrl
'*功能：对控件进行语法高亮显示
'*说明：
'*传入参数：
'*    ctrl           类型：RichTextBox      --控件
'*    filename       类型：String           --临时文件名
'*    bKillFile      类型：Boolean          --完成后是否删除文件
'*返回参数：
'*
'*作者：progame  日期：2002-09-29  17:38:22
'**************************************************************
Private Sub HighLightCtrl(ctrl As RichTextBox, filename As String, bKillFile As Boolean)
Dim sLine() As String
Dim i       As Long
Dim l       As Long
Dim sTmp        As String
Dim iProgress   As Integer      '*当前处理进度
Dim iTmp        As Integer

    
    sTmp = ctrl.Text
    l = Len(sTmp)
    If l = 0 Then
        Exit Sub
    End If
    '*先将
    Open filename For Output As #1
    
        Print #1, HEAD_FIRST & m_sColorRtf & HEAD_LAST
 
        m_iQuotState = 0
        m_iCommState = 0
        
        iProgress = 0
        
        '*对于大文本进行分块处理，每1K为一个单位
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
            '*处理功能字符和分隔字符
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
            
                Call HighLightLine(sLine(j))    '*处理一行
                
                '*恢复分隔字符
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
'*扫描字符，输出格式化后的字符串
Dim i       As Long
Dim j       As Integer
Dim k       As Integer
Dim bComm As Boolean

    m_sWord = Split(sLine, " ")
    sLine = ""
    
    k = UBound(m_sWord)
    
    For i = 0 To k

        '*如果不是字符串，则进行先进行注释判断，再进行关键字判断
        If m_iQuotState = 0 Then
        
            bComm = TreatComment(m_sWord(i), (i = k))         '*处理注释
        
        End If

        If Not bComm Then

                If TreatQuot(m_sWord(i), (i = k), 1) Then       '*处理字符串
                
                    '*恢复被处理过的字符串标志
                    If m_sQuotFlag1 <> "" Then
                        m_sWord(i) = Replace(m_sWord(i), Chr(1), m_sQuotFlag1)
                    End If
                    If m_sQuotFlag2 <> "" Then
                        m_sWord(i) = Replace(m_sWord(i), Chr(2), m_sQuotFlag2)
                    End If
                Else
                
                    If m_iQuotState = 0 Then
                        Call TreatKeyWord(m_sWord(i))         '*处理关键字
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
'*处理注释
'*      bLast       --是否最后一个字符
'*      返回值：    sWord是否注释
    Select Case m_iCommState
        Case 0
            If m_sLnCommFlag1 <> "" Then
                If InStr(1, sWord, m_sLnCommFlag1, vbBinaryCompare) = 1 Then  '*是单行注释一
                    m_iCommState = 1
                    '*加上注释字体RTF格式头
                    sWord = m_sCommRtfHead & sWord
                    TreatComment = True
                    '*如果最后一行，再进行处理
                    If bLast Then
                        TreatComment sWord, True
                    End If
                    Exit Function
                End If
            End If
            If m_sLnCommFlag2 <> "" Then
                If InStr(1, sWord, m_sLnCommFlag2, vbBinaryCompare) = 1 Then  '*是单行注释二
                    m_iCommState = 2
                    '*加上注释字体RTF格式头
                    sWord = m_sCommRtfHead & sWord
                    TreatComment = True
                    '*如果最后一个词，再进行处理
                    If bLast Then
                        TreatComment sWord, True
                    End If
                    Exit Function
                End If
            End If
            If m_sBlkCommHead1 <> "" And m_sBlkCommFoot1 <> "" Then
                If InStr(1, sWord, m_sBlkCommHead1, vbBinaryCompare) = 1 Then  '*是多行注释一
                    m_iCommState = 3
                    '*加上注释字体RTF格式头
                    sWord = m_sCommRtfHead & sWord
                    TreatComment = True
                    TreatComment sWord, bLast
                    Exit Function
                End If
            End If
            If m_sBlkCommHead2 <> "" And m_sBlkCommFoot2 <> "" Then
                If InStr(1, sWord, m_sBlkCommHead2, vbBinaryCompare) = 1 Then '*是多行注释二
                    m_iCommState = 4
                    '*加上注释字体RTF格式头
                    sWord = m_sCommRtfHead & sWord
                    TreatComment = True
                    TreatComment sWord, bLast
                    Exit Function
                End If
            End If
        Case 1, 2
            TreatComment = True
            If bLast Then       '*单行注释，且当前为此行最后一个词，将注释标志恢复为0
                sWord = sWord & m_sCommRtfFoot
                m_iCommState = 0
            End If
        Case 3
            TreatComment = True
            If m_sBlkCommHead1 <> "" And m_sBlkCommFoot1 <> "" And Len(sWord) > 1 Then
                If InStrRev(sWord, m_sBlkCommFoot1, , vbBinaryCompare) <> 0 Then '*是多行注释一
                    m_iCommState = 0
                    '*加上注释字体RTF格式尾
                    sWord = sWord & m_sCommRtfFoot
                    Exit Function
                End If
            End If
        Case 4
            TreatComment = True
            If m_sBlkCommHead2 <> "" And m_sBlkCommFoot2 <> "" And Len(sWord) > 1 Then
                If InStrRev(sWord, m_sBlkCommFoot2, , vbBinaryCompare) <> 0 Then '*是多行注释二
                    m_iCommState = 0
                    '*加上注释字体RTF格式尾
                    sWord = sWord & m_sCommRtfFoot
                    Exit Function
                End If
            End If
        Case Else
        
    End Select
    
End Function

Private Function TreatKeyWord(ByRef sWord As String) As Boolean
'*处理关键字
'*      返回值：    sWord是否关键字
Dim i           As Integer
Dim bKeyWord    As Boolean
    '* 先处理是不是数字
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
            '*加上字体RTF格式的头尾于词两端
            sWord = m_sKwdRtfHead(i) & sWord & m_sKwdRtfFoot(i)
            TreatKeyWord = True
            Exit Function
        End If
    Next i
End Function

Private Function TreatQuot(ByRef sWord As String, bLast As Boolean, iBegin As Integer) As Boolean
'*处理字符串
'*      bLast       是否此行最后一个词
'*      iBegin      搜索位置
'*      返回值：    是否做过字符串处理
Dim i       As Integer
Dim j       As Integer
Dim sQuot   As String       '*当前用来判断的字符串边界符
    Select Case m_iQuotState
        Case 0
            '*判断哪种字符串标志先出现，使用先出现的字符串标志来决定字符串起始位置
            i = 0
            j = 0
            If m_sQuotFlag1 <> "" Then
                i = InStr(iBegin, sWord, m_sQuotFlag1, vbBinaryCompare)
            End If
            If m_sQuotFlag2 <> "" Then
                j = InStr(iBegin, sWord, m_sQuotFlag2, vbBinaryCompare)
            End If
            If i = 0 And j = 0 Then     '*没有字符串标志出现，不用判断
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
    
    '*查找当前的字符串标志
    i = InStr(iBegin, sWord, sQuot, vbBinaryCompare)
    
    '*如果前一个字符是敏感字符前缀，则此次查找无效
    If m_sEscape <> "" Then
        If i > 1 Then
            j = InStr(iBegin, sWord, m_sEscape)
            If j <> 0 And j = i - Len(m_sEscape) Then
                '*将此字符串标志屏避
                If sQuot = m_sQuotFlag1 Then
                    sWord = Replace(sWord, sQuot, Chr(1), , 1)
                Else
                    sWord = Replace(sWord, sQuot, Chr(2), , 1)
                End If
                TreatQuot = True
                '*继续递归处理
                Call TreatQuot(sWord, bLast, i + 1)
                Exit Function
            End If
        End If
    End If
    
    If i <> 0 Then
    
        TreatQuot = True
        
        '*格式化字符串
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

        '*继续递归处理
        Call TreatQuot(sWord, bLast, i)
        
    Else
    
        '*如果已经没有了标志，但续行无效的情况下，当前这个词为最后一个时，强行结束
        If bLast And Not m_bMultiLine And m_iQuotState <> 0 Then
            sWord = sWord & m_sQuotRtfFoot
            m_iQuotState = 0
        End If
        
    End If
    

End Function


'**************************************************************
'*Rtf2Html
'*功能：    将RTF控件内容转换为Html字符串
'*说明：
'*传入参数：
'*    rTxt      类型：RichTextBox
'*返回参数：
'*      HTML字符串
'*作者：progame  日期：2002-10-17  17:36:17
'**************************************************************
Public Function Rtf2Html(ctrl As RichTextBox) As String
Dim byt()       As Byte
Dim sFileName   As String
    sFileName = TempFileName
    
    HighLightCtrl ctrl, sFileName, False
    
    '*读入文件
    Open sFileName For Binary As #1
        ReDim byt(1 To FileLen(sFileName))
        Get #1, , byt
    Close #1
    
    If Dir(sFileName) <> "" Then
        Kill sFileName
    End If
    
    Rtf2Html = StrConv(byt, vbUnicode)
    
    '*将特殊字符恢复
    Rtf2Html = Replace(Rtf2Html, "\\", "\")
    Rtf2Html = Replace(Rtf2Html, "\{", "{")
    Rtf2Html = Replace(Rtf2Html, "\}", "}")
    
    '*先删除RTF文件头
    Rtf2Html = Replace(Rtf2Html, HEAD_FIRST & m_sColorRtf & HEAD_LAST, "", , 1)
    '*转换换行符号
    Rtf2Html = Replace(Rtf2Html, "\par ", "")
    
    '*将特殊字符处理
    Rtf2Html = Replace(Rtf2Html, "&", "&amp;")
    Rtf2Html = Replace(Rtf2Html, "<", "&lt;")
    Rtf2Html = Replace(Rtf2Html, ">", "&gt;")
    '*转换注释字体符号
    FontRtf2Html Rtf2Html, m_objQuotFont
    FontRtf2Html Rtf2Html, m_objCommFont
    Dim i       As Integer
    For i = 1 To m_iKwdCount
        FontRtf2Html Rtf2Html, m_objKwdFont(i)
    Next i
    
    '*去除最后一个"}"
    Rtf2Html = Left(Rtf2Html, Len(Rtf2Html) - 4)
    
    
    '*加上HTML头尾
    Rtf2Html = HEAD_HTML & Rtf2Html & TAIL_HTML
 
End Function

Private Sub FontRtf2Html(ByRef sDeal As String, ByRef objFont As CFont)
'*对sDeal字符串进行RTF到HTML的字体转换
    sDeal = Replace(sDeal, objFont.getRtfHead(m_objColorManager), objFont.getHtmlHead(m_objColorManager))
    sDeal = Replace(sDeal, objFont.getRtfTail, objFont.getHtmlTail)
End Sub

'**************************************************************
'*Rtf2Ubb
'*功能：    将RTF控件内容转换为Ubb字符串
'*说明：
'*传入参数：
'*    rTxt      类型：RichTextBox
'*返回参数：
'*      HTML字符串
'*作者：progame  日期：2002-10-17  17:36:17
'**************************************************************
Public Function Rtf2Ubb(ctrl As RichTextBox) As String
Dim byt()       As Byte
Dim sFileName   As String
    sFileName = TempFileName
    
    HighLightCtrl ctrl, sFileName, False

    '*读入文件
    Open sFileName For Binary As #1
        ReDim byt(1 To FileLen(sFileName))
        Get #1, , byt
    Close #1
    
    If Dir(sFileName) <> "" Then
        Kill sFileName
    End If
    
    Rtf2Ubb = StrConv(byt, vbUnicode)
    
    '*将特殊字符恢复
    Rtf2Ubb = Replace(Rtf2Ubb, "\\", "\")
    Rtf2Ubb = Replace(Rtf2Ubb, "\{", "{")
    Rtf2Ubb = Replace(Rtf2Ubb, "\}", "}")
    
    '*先删除RTF文件头
    Rtf2Ubb = Replace(Rtf2Ubb, HEAD_FIRST & m_sColorRtf & HEAD_LAST, "", , 1)
    '*转换换行符号
    Rtf2Ubb = Replace(Rtf2Ubb, "\par ", vbCrLf)
    
    '*转换注释字体符号
    FontRtf2Ubb Rtf2Ubb, m_objQuotFont
    FontRtf2Ubb Rtf2Ubb, m_objCommFont
    Dim i       As Integer
    For i = 1 To m_iKwdCount
        FontRtf2Ubb Rtf2Ubb, m_objKwdFont(i)
    Next i
    
    '*去除最后一个"}"
    Rtf2Ubb = Left(Rtf2Ubb, Len(Rtf2Ubb) - 4)
    
 
End Function

Private Sub FontRtf2Ubb(ByRef sDeal As String, ByRef objFont As CFont)
'*对sDeal字符串进行RTF到UBB的字体转换
    sDeal = Replace(sDeal, objFont.getRtfHead(m_objColorManager), objFont.getUBBHead(m_objColorManager))
    sDeal = Replace(sDeal, objFont.getRtfTail, objFont.getUBBTail)
End Sub
