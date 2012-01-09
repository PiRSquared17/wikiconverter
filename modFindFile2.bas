Attribute VB_Name = "modFindFile"
Public Const MAX_PATH = 260
Public Const INVALID_HANDLE_VALUE = -1
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
'Download by http://www.codefans.net
Public Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type

Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Public Sub FindFiles(strRootFolder As String, strFolder As String, strFile As String)
    
    Dim lngSearchHandle As Long
    Dim udtFindData As WIN32_FIND_DATA
    Dim strTemp As String, lngRet As Long
        
    '����ļ����Ƿ��� \
    If Right$(strRootFolder, 1) <> "\" Then strRootFolder = strRootFolder & "\"
    
    '������һ���ļ����
    lngSearchHandle = FindFirstFile(strRootFolder & "*", udtFindData)
    
    '�����Чʱ�˳�
    If lngSearchHandle = INVALID_HANDLE_VALUE Then Exit Sub
    
    lngRet = 1
    
    Do While lngRet <> 0
        
        'ȥ���ո�
        strTemp = TrimNulls(udtFindData.cFileName)
        
        If (udtFindData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
            'ΪĿ¼ʱ
            '������.��..ʱ
            If strTemp <> "." And strTemp <> ".." Then
                'frmMain.List2.AddItem "Ŀ¼ | " & strTemp
                Call FindFiles(strRootFolder & strTemp, strFolder, strFile)
            End If
        Else
            'Ϊ�ļ�ʱ
            'frmMain.List2.AddItem "�ļ� | " & strTemp
            'ΪTxt�ļ�ʱ
            If Right$(strTemp, 4) = ".txt" Then
                frmMain.List2.AddItem "�ļ� | " & strTemp
                'Call AppendWikiXML(strRootFolder, strTemp)
                'Call AppendMWikiXML(ҳ����⣬ҳ��ID��ҳ�����ݣ��汾ID)
                'Call AppendMWikiXML("strPageTitle", 1, "strPageContent", 1)
            End If
        End If
        
        '������һ���ļ���Ŀ¼���
        lngRet = FindNextFile(lngSearchHandle, udtFindData)
        
    Loop
    
    Call FindClose(lngSearchHandle)
    
End Sub

Public Function TrimNulls(strString As String) As String
   
   Dim l As Long
   
   l = InStr(1, strString, Chr(0))
   
   If l = 1 Then
      TrimNulls = ""
   ElseIf l > 0 Then
      TrimNulls = Left$(strString, l - 1)
   Else
      TrimNulls = strString
   End If
   
End Function

