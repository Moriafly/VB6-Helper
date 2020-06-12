Attribute VB_Name = "ModFun"
'**************************************************************
'*模块名称：ModFun
'*模块功能：
'*说明：http://www.NewXing.com
'*作者：progame  2002-09-29  13:02:24
'***************************************************************
Option Explicit


'*RTF文件头(First和Last之间插入颜色和字体信息
Public Const HEAD_FIRST = "{\rtf1\ansi\ansicpg936\deff0{\fonttbl}{\colortbl ;"
Public Const HEAD_LAST = "}\viewkind4\uc1\pard\lang2052\f0\fs18 "

Public Const HEAD_HTML = "<PRE>"
Public Const TAIL_HTML = "</PRE>"

Private Declare Function GetTempFileName _
    Lib "kernel32" Alias "GetTempFileNameA" _
                (ByVal lpszPath As String, _
                ByVal lpPrefixString As String, _
                ByVal wUnique As Long, _
                ByVal lpTempFileName As String) _
    As Long
Private Declare Function GetTempPath _
    Lib "kernel32" Alias "GetTempPathA" ( _
                ByVal nBufferLength As Long, _
                ByVal lpBuffer As String) _
    As Long
    
Private Const MAX_PATH = 255


Public Function TempFileName() As String
'*取得临时文件名
Dim temp_path As String
Dim temp_file As String
Dim length As Long

    temp_path = Space(MAX_PATH)
    length = GetTempPath(MAX_PATH, temp_path)
    temp_path = Left(temp_path, length)
    
    temp_file = Space(MAX_PATH)
    GetTempFileName temp_path, "per", 0, temp_file
    TempFileName = Left(temp_file, InStr(temp_file, Chr(0)) - 1)
    
End Function

