VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BatchEncryption DeCoder"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9360
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   9360
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdDeCodeAndOutputToFile 
      Caption         =   "解密并输出"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   6120
      Width           =   9135
   End
   Begin VB.CommandButton cmdOutFile 
      Caption         =   "..."
      Height          =   375
      Left            =   8520
      TabIndex        =   4
      ToolTipText     =   "选择输出文件"
      Top             =   5640
      Width           =   735
   End
   Begin VB.CommandButton cmdInFile 
      Caption         =   "..."
      Height          =   375
      Left            =   8520
      TabIndex        =   1
      ToolTipText     =   "选择待解密文件"
      Top             =   5160
      Width           =   735
   End
   Begin SHDocVwCtl.WebBrowser brwDocument 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      ExtentX         =   16113
      ExtentY         =   8705
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label lblOutFile 
      Caption         =   "输出文件："
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   5685
      Width           =   8295
   End
   Begin VB.Label lblInFile 
      Caption         =   "待解密文件："
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   5205
      Width           =   8295
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim blnCanWriteData As Boolean
Dim blnIsTitleShowed As Boolean

Dim strInFile As String
Dim strOutFile As String

Dim strCode As String
Dim strEnvironmentList() As String

Dim albAlphaBet As ALPHABET

Dim cmsPasswordTable As CHRMAPSET

Private Sub brwDocument_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    blnCanWriteData = True
    If blnIsTitleShowed = False Then
        With brwDocument.Document
            .Open
            .Clear
            .Write StrConv(LoadResData(101, 23), vbUnicode)
            '.Close
        End With
        blnIsTitleShowed = True
    End If
End Sub

Private Sub cmdDeCodeAndOutputToFile_Click()

    If strInFile = "" Then
        MsgBox "无效文件设置，没有选定源文件", vbCritical, App.Title
        Exit Sub
    End If

    If strOutFile = "" Then
        MsgBox "无效文件设置，没有选定目标文件", vbCritical, App.Title
        Exit Sub
    End If

    If UCase(strInFile) = UCase(strOutFile) Then
        MsgBox "无效文件设置，源文件名不得与目标文件名相同", vbCritical, App.Title
        Exit Sub
    End If
    DoEvents
    DeCodeAndOutput
End Sub

Private Sub cmdInFile_Click()
    Dim dlrReturn As DLGRET
    dlrReturn = GetOpenFile(Me.hWnd, "MS-DOS 批处理文件 (*.bat)" & Chr(0) & "*.bat" & Chr(0) & _
                                    "Windows NT 批处理文件 (*.cmd)" & Chr(0) & "*.cmd" & Chr(0), _
                                    App.Path, "选择待解密文件")
    If dlrReturn.blnIsOpened = True Then
        strInFile = dlrReturn.strFileName
        lblInFile.Caption = "待解密文件：" & strInFile
        ReadInFile
        WriteInFileInformation
    Else
        MsgBox "无法访问指定的文件", vbCritical, App.Title
    End If
End Sub

Private Sub cmdOutFile_Click()
    Dim dlrReturn As DLGRET
    dlrReturn = GetSaveFile(Me.hWnd, "MS-DOS 批处理文件 (*.bat)" & Chr(0) & "*.bat" & Chr(0) & _
                                    "Windows NT 批处理文件 (*.cmd)" & Chr(0) & "*.cmd" & Chr(0), _
                                    App.Path, "选择输出文件")
                                    
    If dlrReturn.blnIsOpened = True Then
        strOutFile = dlrReturn.strFileName
        lblOutFile.Caption = "输出文件：" & strOutFile
    Else
        MsgBox "无法访问指定的文件", vbCritical, App.Title
    End If
End Sub

Private Sub Form_Initialize()
    blnCanWriteData = False
    blnIsTitleShowed = False
    brwDocument.Navigate "about:blank"
End Sub

Private Sub Form_Load()
    '
End Sub

Private Function ReadInFile()
    Dim intFileNum As Integer
    Dim strNextLine As String

    strCode = ""
    
    intFileNum = FreeFile

    Open strInFile For Input As #intFileNum
    Do Until EOF(intFileNum)
        DoEvents
        Line Input #intFileNum, strNextLine
        strCode = strCode & strNextLine & vbCrLf
    Loop
    Close #intFileNum
    
    strCode = Mid(strCode, 3)
    
End Function

Private Function WriteInFileInformation()
    With brwDocument.Document
        '.Open
        .Write "<div class=""sectionTitle""><h5>待解密文件：" & ChangeTextToHTMLEntity(strInFile)
        .Write "<hr />代码</h5>" & vbCrLf
        '.Close
        
        .Write "<div class=""programCode"">"
        .Write "<p><h6>"
        DoEvents
        .Write Replace(ChangeTextToHTMLEntity(strCode), vbCrLf, "<br />")
        DoEvents
        .Write "</h6></p></div></div>" & vbCrLf
        
    End With
End Function

Private Function WriteEncryptionHeader(strEncryptionHeader As String)
    With brwDocument.Document
        '.Open
        .Write "<div class=""sectionTitle""><h5>加密标头</h5>" & vbCrLf
        '.Close
        
        .Write "<div class=""programCode"">"
        .Write "<p><h6>"
        DoEvents
        .Write Replace(ChangeTextToHTMLEntity(strEncryptionHeader), vbCrLf, "<br />")
        DoEvents
        .Write "</h6></p></div></div>" & vbCrLf
        
    End With
End Function

Private Function WriteEnvironmentList()

    Dim i As Long

    With brwDocument.Document
        '.Open
        .Write "<div class=""sectionTitle""><h5>加密标头所使用的环境变量</h5>" & vbCrLf
        '.Close
        
        .Write "<div class=""programCode"">"
        .Write "<p><h6>"
        For i = 0 To UBound(strEnvironmentList)
            DoEvents
            .Write ChangeTextToHTMLEntity(strEnvironmentList(i))
            .Write "<br />"
            DoEvents
        Next i
        .Write "</h6></p></div></div>" & vbCrLf
        
    End With
End Function

Private Function WriteAlphaBetList()

    Dim i

    With brwDocument.Document
        '.Open
        .Write "<div class=""sectionTitle""><h5>环境变量映射表</h5>" & vbCrLf
        '.Close
        
        .Write "<div class=""programCode"">"
        .Write "<p><h6>" & vbCrLf
        .Write "<table border=""1"" style=""font-size: xx-small;""><tr><td>环境变量映射</td><td>映射内容</td><tr>" & vbCrLf
        For i = 0 To UBound(albAlphaBet.cmsCharMapSet)
            For j = 0 To UBound(albAlphaBet.cmsCharMapSet(i).cmpCharMap)
                DoEvents
                .Write "<tr><td>"
                .Write ChangeTextToHTMLEntity(albAlphaBet.cmsCharMapSet(i).cmpCharMap(j).strCipherText)
                .Write "</td><td>"
                .Write ChangeTextToHTMLEntity(albAlphaBet.cmsCharMapSet(i).cmpCharMap(j).strPlainText)
                .Write "</tr>"
                DoEvents
            Next j
        Next i
        .Write "</table></h6></p></div></div>" & vbCrLf
        
    End With
End Function

Private Function WriteHeader(strHeader As String)
    With brwDocument.Document
        '.Open
        .Write "<div class=""sectionTitle""><h5>加密标头</h5>" & vbCrLf
        '.Close
        
        .Write "<div class=""programCode"">"
        .Write "<p><h6>"
        DoEvents
        .Write ChangeTextToHTMLEntity(strHeader)
        DoEvents
        .Write "</h6></p></div></div>" & vbCrLf
        
    End With
End Function

Private Function WritePasswordTable()
    With brwDocument.Document
        '.Open
        .Write "<div class=""sectionTitle""><h5>密码表</h5>" & vbCrLf
        '.Close
        
        .Write "<div class=""programCode"">"
        .Write "<p><h6>"
        DoEvents
        .Write ChangeTextToHTMLEntity(cmsPasswordTable.strEnvironmentValue)
        DoEvents
        .Write "</h6></p></div></div>" & vbCrLf
        
    End With
End Function

Private Function WriteCode(strSrcCode As String, lngCount As Long)
    With brwDocument.Document
        '.Open
        .Write "<div class=""sectionTitle""><h5>第 " & lngCount & " 轮解密</h5>" & vbCrLf
        '.Close
        
        .Write "<div class=""programCode"">"
        .Write "<p><h6>"
        DoEvents
        .Write Replace(ChangeTextToHTMLEntity(strSrcCode), vbCrLf, "<br />")
        DoEvents
        .Write "</h6></p></div></div>" & vbCrLf
        
    End With
End Function

Private Function WriteSrcCode(strSrcCode As String)
    With brwDocument.Document
        '.Open
        .Write "<div class=""sectionTitle""><h5>源码</h5>" & vbCrLf
        '.Close
        
        .Write "<div class=""programCode"">"
        .Write "<p><h6>"
        DoEvents
        .Write Replace(ChangeTextToHTMLEntity(strSrcCode), vbCrLf, "<br />")
        DoEvents
        .Write "</h6></p></div></div>" & vbCrLf
        
    End With
End Function

' ------------------------------------------------------------------------------------------------
' 解密函数
' ------------------------------------------------------------------------------------------------

Private Function DeCodeAndOutput()
    Dim strEncryptionHeader As String
    Dim strHeader As String
    Dim strSourceCode As String ' 源码
    
    Me.Caption = "BatchEncryption DeCoder [Working...]"
    
    strEncryptionHeader = GetEncryptionHeader()
    WriteEncryptionHeader strEncryptionHeader
    
    GetEnvironmentList strEncryptionHeader
    WriteEnvironmentList
    
    InitAlphaBet
    WriteAlphaBetList
    
    strHeader = GetHeader(strEncryptionHeader)
    WriteHeader strHeader
    
    'GetPasswordTable strHeader
    'WritePasswordTable
    
    strSourceCode = DeCode
    Output strSourceCode
    
    Me.Caption = "BatchEncryption DeCoder"
    
    MsgBox "解密成功！", vbInformation, App.Title
    
End Function

' 输出到文件
Private Function Output(strSrcCode As String)
    Dim intFileNum As Integer
    intFileNum = FreeFile
    
    Me.Caption = "BatchEncryption DeCoder [Working...] [输出到文件...]"
    
    Open strOutFile For Output As #intFileNum
        Print #intFileNum, strSrcCode
    Close #intFileNum
End Function

' 解密
Private Function DeCode() As String
    Dim strSrc As String ' 需处理的源码
    Dim strTemp As String ' 临时变量
    Dim strFirstLine As String ' 第一行
    Dim lngCount As Long ' 解密轮数
    Dim i As Long
    Dim j As Long
    
    ' 思路：
    ' 去掉开头多余行后，对每一行进行分析，如果是密码表则进行分析，如果不是就开始全文解密
    
    strSrc = strCode
    
    ' 去前 2 行
    For i = 0 To 1
        strSrc = Mid(strSrc, InStr(strSrc, vbCrLf) + 2)
    Next i
    
    ' Clipboard.SetText strSrc
    strFirstLine = Mid(strSrc, 1, InStr(strSrc, vbCrLf) - 1)
    ' Clipboard.SetText strFirstLine
    
    ' 格式化代码
    strFirstLine = GetHeader(strFirstLine)
    strSrc = Mid(strSrc, InStr(strSrc, vbCrLf) + 2)
    strSrc = strFirstLine & vbCrLf & strSrc
    ' Clipboard.SetText strSrc
    
    lngCount = 1
    
    Do While UCase(strFirstLine) Like UCase("*@set '=*") ' 若第一行有密码表设置语句
    
        Me.Caption = "BatchEncryption DeCoder [Working...] [第 " & lngCount & " 轮解密]"
    
        ' 获取密码
        GetPasswordTable strFirstLine
        WritePasswordTable
        InitCharMapSet cmsPasswordTable
        
        strSrc = Mid(strSrc, InStr(strSrc, vbCrLf) + 2) ' 去第一行
        strTemp = strSrc
        
        ' 全文解密
        For i = 0 To UBound(cmsPasswordTable.cmpCharMap)
            DoEvents
            With cmsPasswordTable.cmpCharMap(i)
                strTemp = Replace(strTemp, "%" & .strCipherText & "%", .strPlainText)
            End With
        Next i
        
        ' 获取下一行
        strFirstLine = Mid(strTemp, 1, InStr(strTemp, vbCrLf) - 1)
        
        ' 输出到 brwDocument
        WriteCode strTemp, lngCount
        
        lngCount = lngCount + 1
    
    Loop
    
    strTemp = Mid(strTemp, InStr(strTemp, vbCrLf) + 2) ' 去第一行（还原行）
    WriteSrcCode strTemp
    
    DeCode = strTemp ' 返回
    
End Function

' 获取密码表
Private Function GetPasswordTable(strSrc As String)
    Dim strTemp As String
    Dim vntTemp As Variant
    Dim strPassword As String
    Dim i As Long
    
    strTemp = strSrc
    strTemp = Replace(strTemp, "^^", "转")
    strTemp = Replace(strTemp, "^&", "与")
    strTemp = Replace(strTemp, "^", "")
    
    strTemp = Replace(strTemp, "&", vbCrLf)
    
    ' MsgBox strTemp
    
    vntTemp = Split(strTemp, vbCrLf)
    For i = 0 To UBound(vntTemp)
        If UCase(vntTemp(i)) Like UCase("@set '=*") Then
            strPassword = vntTemp(i)
        End If
    Next i
    
    strPassword = Replace(strPassword, "转", "^")
    strPassword = Replace(strPassword, "与", "&")
    
    strPassword = Mid(strPassword, 8)
    
    cmsPasswordTable.strEnvironmentName = "'"
    cmsPasswordTable.strEnvironmentValue = strPassword
    
End Function

' 获取标头
Private Function GetHeader(strEncryptionHeader As String)
    Dim i As Long
    Dim j As Long
    Dim strTemp As String
    
    strTemp = strEncryptionHeader
    
    For i = 0 To UBound(albAlphaBet.cmsCharMapSet)
        For j = 0 To UBound(albAlphaBet.cmsCharMapSet(i).cmpCharMap)
            With albAlphaBet.cmsCharMapSet(i).cmpCharMap(j)
                ' MsgBox i & vbCrLf & j & vbCrLf & .strCipherText & vbCrLf & .strPlainText
                strTemp = Replace(strTemp, "%" & .strCipherText & "%", .strPlainText)
            End With
        Next j
    Next i
    
    'MsgBox strTemp
    
    GetHeader = strTemp
    
End Function

' 初始化字母表
Private Function InitAlphaBet()
    Dim i As Long
    
    ReDim albAlphaBet.cmsCharMapSet(UBound(strEnvironmentList))
    
    For i = 0 To UBound(strEnvironmentList)
        albAlphaBet.cmsCharMapSet(i).strEnvironmentName = strEnvironmentList(i)
        albAlphaBet.cmsCharMapSet(i).strEnvironmentValue = Replace(Environ(strEnvironmentList(i)), "Program Files (x86)", "Program Files")
        InitCharMapSet albAlphaBet.cmsCharMapSet(i)
    Next i
    
End Function

' 初始化映射字符集
Private Function InitCharMapSet(ByRef cmsCharMapSet As CHRMAPSET)
    Dim lngEnvironmentSize As Long
    Dim lngMapCount As Long
    Dim lngMapIndex As Long
    Dim i As Long
    Dim j As Long
    
    lngEnvironmentSize = Len(cmsCharMapSet.strEnvironmentValue)
    
    ' %strEnvName:~lngOffset,lngLength% 类
    lngMapCount = ((lngEnvironmentSize ^ 2) + lngEnvironmentSize) / 2
    ReDim cmsCharMapSet.cmpCharMap(lngMapCount)
    
    With cmsCharMapSet.cmpCharMap(0)
        .strCipherText = cmsCharMapSet.strEnvironmentName
        .strPlainText = cmsCharMapSet.strEnvironmentValue
    End With
    
    lngMapIndex = 1
    
    For i = 0 To lngEnvironmentSize - 1 ' 起始字符
        For j = 1 To lngEnvironmentSize - i ' 长度
            With cmsCharMapSet.cmpCharMap(lngMapIndex)
                .strCipherText = cmsCharMapSet.strEnvironmentName & ":~" & i & "," & j
                .strPlainText = Mid(cmsCharMapSet.strEnvironmentValue, i + 1, j)
            End With
            lngMapIndex = lngMapIndex + 1
        Next j
    Next i
    
    ' %strEnvName:~-lngLength% 类
    lngMapCount = lngMapCount + lngEnvironmentSize
    ReDim Preserve cmsCharMapSet.cmpCharMap(lngMapCount)
    For i = lngEnvironmentSize To 1 Step -1
        With cmsCharMapSet.cmpCharMap(lngMapIndex)
            .strCipherText = cmsCharMapSet.strEnvironmentName & ":~-" & i
            .strPlainText = Mid(cmsCharMapSet.strEnvironmentValue, lngEnvironmentSize - i + 1, i)
        End With
        lngMapIndex = lngMapIndex + 1
    Next i
    
    ' %strEnvName:~lngStart,-lngEnd% 类
    lngMapCount = lngMapCount + (((lngEnvironmentSize ^ 2) - lngEnvironmentSize) / 2)
    ReDim Preserve cmsCharMapSet.cmpCharMap(lngMapCount)
    For i = 0 To lngEnvironmentSize - 2 ' 起始字符
        For j = lngEnvironmentSize - i - 1 To 1 Step -1 ' 结束字符
            With cmsCharMapSet.cmpCharMap(lngMapIndex)
                .strCipherText = cmsCharMapSet.strEnvironmentName & ":~" & i & ",-" & j
                .strPlainText = Mid(cmsCharMapSet.strEnvironmentValue, i + 1, lngEnvironmentSize - i - j)
            End With
            lngMapIndex = lngMapIndex + 1
        Next j
    Next i
    
    ' %strEnvName:~-lngStart,-lngEnd% 类
    lngMapCount = lngMapCount + (((lngEnvironmentSize ^ 2) - lngEnvironmentSize) / 2)
    ReDim Preserve cmsCharMapSet.cmpCharMap(lngMapCount)
    For i = lngEnvironmentSize To 2 Step -1 ' 起始字符
        For j = i - 1 To 1 Step -1 ' 结束字符
            With cmsCharMapSet.cmpCharMap(lngMapIndex)
                .strCipherText = cmsCharMapSet.strEnvironmentName & ":~-" & i & ",-" & j
                .strPlainText = Mid(cmsCharMapSet.strEnvironmentValue, lngEnvironmentSize - i + 1, i - j)
            End With
            lngMapIndex = lngMapIndex + 1
        Next j
    Next i
    
    ' %strEnvName:~-lngStart,lngLength% 类
    lngMapCount = lngMapCount + (((lngEnvironmentSize ^ 2) + lngEnvironmentSize) / 2)
    ReDim Preserve cmsCharMapSet.cmpCharMap(lngMapCount)
    For i = lngEnvironmentSize To 1 Step -1 ' 起始字符
        For j = i To 1 Step -1 ' 长度
            With cmsCharMapSet.cmpCharMap(lngMapIndex)
                .strCipherText = cmsCharMapSet.strEnvironmentName & ":~-" & i & "," & j
                .strPlainText = Mid(cmsCharMapSet.strEnvironmentValue, lngEnvironmentSize - i + 1, j)
            End With
            lngMapIndex = lngMapIndex + 1
        Next j
    Next i
    
    ' If idxIndex > 0 Then Exit Function
    ' For i = 0 To lngMapCount
        ' MsgBox i & vbCrLf & _
                cmsCharMapSet.cmpCharMap(i).strCipherText & vbCrLf & _
                cmsCharMapSet.cmpCharMap(i).strPlainText
    ' Next i
    
End Function

' 获取环境变量列表
Private Function GetEnvironmentList(strEncryptionHeader As String)
    Dim strTemp As String
    Dim vntTemp As Variant
    Dim strList As String
    Dim vntList As Variant
    Dim i As Long
    Dim j As Long
    
    strTemp = strEncryptionHeader
    
    For i = 0 To 9
        strTemp = Replace(strTemp, i, "")
    Next i
    
    strTemp = Replace(strTemp, "%", ",")
    strTemp = Replace(strTemp, "^", "")
    strTemp = Replace(strTemp, "-", "")
    strTemp = Replace(strTemp, ":", "")
    strTemp = Replace(strTemp, "~", "")
    
    strTemp = Replace(strTemp, "'", "")
    strTemp = Replace(strTemp, "=", "")
    strTemp = Replace(strTemp, "@", "")
    strTemp = Replace(strTemp, "&", "")
    
    strTemp = Replace(strTemp, "#", "")
    strTemp = Replace(strTemp, "\", "")
    strTemp = Replace(strTemp, ">", "")
    
    strTemp = Replace(strTemp, "+", "")
    strTemp = Replace(strTemp, "$", "")
    strTemp = Replace(strTemp, ")", "")
    strTemp = Replace(strTemp, "(", "")
    strTemp = Replace(strTemp, "[", "")
    strTemp = Replace(strTemp, "]", "")
    strTemp = Replace(strTemp, "{", "")
    strTemp = Replace(strTemp, "}", "")
    
    strTemp = Replace(strTemp, "<", "")
    strTemp = Replace(strTemp, "?", "")
    
    strTemp = Replace(strTemp, "*", "")
    strTemp = Replace(strTemp, """", "")
    
    strTemp = Replace(strTemp, "_", "")
    strTemp = Replace(strTemp, "/", "")
    
    strTemp = Replace(strTemp, "`", "")
    strTemp = Replace(strTemp, ";", "")
    
    strTemp = Replace(strTemp, "|", "")
    strTemp = Replace(strTemp, ".", "")
    strTemp = Replace(strTemp, " ", "")
    
    vntTemp = Split(strTemp, ",")
    
    Clipboard.SetText strTemp
    
    strList = vbCrLf
    
    For i = 0 To UBound(vntTemp)
        ' MsgBox vntTemp(i) & ":" & Len(vntTemp(i))
        If CStr(vntTemp(i)) <> "" Then
            If InStr(strList, vbCrLf & vntTemp(i)) = 0 Then
                If Environ(vntTemp(i)) <> "" Then
                    strList = strList & vntTemp(i) & vbCrLf
                End If
            End If
        End If
    Next i
    
    ' MsgBox strList
    
    vntList = strList
    
    vntList = Split(vntList, vbCrLf)
    
    ' MsgBox UBound(vntList)
    
    ReDim strEnvironmentList(UBound(vntList) - 2)
    j = 0
    For i = 0 To UBound(vntList)
        If vntList(i) <> "" Then
            strEnvironmentList(j) = vntList(i)
            j = j + 1
        End If
    Next i
    
End Function

' 获取加密头
Private Function GetEncryptionHeader() As String
    Dim strHeader As String
    Dim vntCode As Variant
    vntCode = strCode
    vntCode = Split(vntCode, vbCrLf)
    GetEncryptionHeader = vntCode(2)
End Function

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
