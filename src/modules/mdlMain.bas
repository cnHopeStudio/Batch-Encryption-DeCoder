Attribute VB_Name = "mdlMain"
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Type DLGRET ' �Զ��巵������
    lngIsOpened As Long
    strFileName As String
    blnIsOpened As Boolean
End Type

Public Type CHRMAP  ' �ı�ӳ��
    strPlainText As String
    strCipherText As String
End Type

Public Type CHRMAPSET   ' ӳ���ַ���
    strEnvironmentName As String
    strEnvironmentValue As String
    cmpCharMap() As CHRMAP
End Type

Public Type ALPHABET    ' ������ĸ��
    cmsCharMapSet() As CHRMAPSET
End Type

Public Const GWL_STYLE = (-16)
Public Const WS_THICKFRAME = &H40000

' lpstrInitialDir ��ʼ��ַ
Public Function GetOpenFile(hwndOwner As Long, lpstrFilter As String, lpstrInitialDir As String, lpstrTitle As String) As DLGRET

    Dim ofnOpenFileName As OPENFILENAME
    Dim dlrReturn As DLGRET
    
    With ofnOpenFileName
        .hwndOwner = hwndOwner
        .hInstance = App.hInstance
        .lpstrFilter = lpstrFilter
        .lpstrFile = Space(&HFE)
        .nMaxFile = &HFF
        .lpstrFileTitle = Space(&HFE)
        .nMaxFileTitle = &HFF
        .lpstrInitialDir = lpstrInitialDir
        .lpstrTitle = lpstrTitle
        .flags = &H1804
        .lStructSize = Len(ofnOpenFileName)
    End With
    
    dlrReturn.lngIsOpened = GetOpenFileName(ofnOpenFileName)
    If dlrReturn.lngIsOpened >= 1 Then
        dlrReturn.strFileName = ofnOpenFileName.lpstrFile
        dlrReturn.blnIsOpened = True
    Else
        dlrReturn.strFileName = vbNullString
        dlrReturn.blnIsOpened = False
    End If
    
    GetOpenFile = dlrReturn
    
End Function

' lpstrInitialDir ��ʼ��ַ
Public Function GetSaveFile(hwndOwner As Long, lpstrFilter As String, lpstrInitialDir As String, lpstrTitle As String) As DLGRET

    Dim ofnSaveFileName As OPENFILENAME
    Dim dlrReturn As DLGRET
    
    With ofnSaveFileName
        .hwndOwner = hwndOwner
        .hInstance = App.hInstance
        .lpstrFilter = lpstrFilter
        .lpstrFile = Space(&HFE)
        .nMaxFile = &HFF
        .lpstrFileTitle = Space(&HFE)
        .nMaxFileTitle = &HFF
        .lpstrInitialDir = lpstrInitialDir
        .lpstrTitle = lpstrTitle
        .flags = &H1804
        .lStructSize = Len(ofnSaveFileName)
    End With
    
    dlrReturn.lngIsOpened = GetSaveFileName(ofnSaveFileName)
    If dlrReturn.lngIsOpened >= 1 Then
        dlrReturn.strFileName = ofnSaveFileName.lpstrFile
        dlrReturn.blnIsOpened = True
    Else
        dlrReturn.strFileName = vbNullString
        dlrReturn.blnIsOpened = False
    End If
    
    GetSaveFile = dlrReturn
    
End Function

Public Function ChangeTextToHTMLEntity(strInString As String) As String
    Dim strOutString As String
    
    strOutString = strInString
    
    strOutString = Replace(strOutString, "&", "&amp;")
    
    strOutString = Replace(strOutString, " ", "&nbsp;")
    strOutString = Replace(strOutString, vbTab, "&emsp;")
    
    strOutString = Replace(strOutString, "<", "&lt;")
    strOutString = Replace(strOutString, ">", "&gt;")
    strOutString = Replace(strOutString, """", "&quot;")
    strOutString = Replace(strOutString, "'", "&#39;")
    
    strOutString = Replace(strOutString, ">", "&gt;")
    strOutString = Replace(strOutString, """", "&quot;")
    strOutString = Replace(strOutString, "'", "&#39;")
    
    ChangeTextToHTMLEntity = strOutString
    
End Function

Sub Main()
    Dim fMain As New frmMain
    fMain.Show
End Sub
