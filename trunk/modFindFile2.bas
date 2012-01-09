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
        
    '检测文件夹是否有 \
    If Right$(strRootFolder, 1) <> "\" Then strRootFolder = strRootFolder & "\"
    
    '给出第一个文件句柄
    lngSearchHandle = FindFirstFile(strRootFolder & "*", udtFindData)
    
    '如果无效时退出
    If lngSearchHandle = INVALID_HANDLE_VALUE Then Exit Sub
    
    lngRet = 1
    
    Do While lngRet <> 0
        
        '去掉空格
        strTemp = TrimNulls(udtFindData.cFileName)
        
        If (udtFindData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
            '为目录时
            '不等于.或..时
            If strTemp <> "." And strTemp <> ".." Then
                'frmMain.List2.AddItem "目录 | " & strTemp
                Call FindFiles(strRootFolder & strTemp, strFolder, strFile)
            End If
        Else
            '为文件时
            'frmMain.List2.AddItem "文件 | " & strTemp
            '为Txt文件时
            If Right$(strTemp, 4) = ".txt" Then
                frmMain.List2.AddItem "文件 | " & strTemp
                'Call AppendWikiXML(strRootFolder, strTemp)
                'Call AppendMWikiXML(页面标题，页面ID，页面内容，版本ID)
                'Call AppendMWikiXML("strPageTitle", 1, "strPageContent", 1)
            End If
        End If
        
        '给出下一个文件或目录句柄
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

