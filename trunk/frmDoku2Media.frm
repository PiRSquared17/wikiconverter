VERSION 5.00
Begin VB.Form frmDoku2Media 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "frmDoku2Media"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'===================================================================================================
'功能:最快速度遍历指定文件夹
'用法:MsgBox EnumFile("C:\", List1)
'解释:将C:\下的所有文件显示到List1中  并弹出总文件数目
'深度:NtQueryDirectoryFile是VB里面DIR以及FindFirstFile,FindNextFile等API的内核实现函数.
'     本来开放给3环使用前,是经过封装了的,自己实现了3环的封装.
'===================================================================================================
Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type
Private Type UNICODE_STRING
    uLength As Integer
    uMaximumLength As Integer
    pBuffer As Long
End Type
Private Type IO_STATUS_BLOCK
    Status As Long
    uInformation As Long
End Type
Private Type OBJECT_ATTRIBUTES
    Length As Long
    RootDirectory As Long
    ObjectName As Long
    Attributes As Long
    SecurityDescriptor As Long
    SecurityQualityOfService As Long
End Type
Private Type FILE_BOTH_DIRECTORY_INFORMATION
    NextEntryOffset As Long
    FileIndex As Long
    CreationTime As LARGE_INTEGER
    LastAccessTime As LARGE_INTEGER
    LastWriteTime As LARGE_INTEGER
    ChangeTime As LARGE_INTEGER
    EndOfFile As LARGE_INTEGER
    AllocationSize As LARGE_INTEGER
    FileAttributes As Long
    FileNameLength As Long
    EaSize As Long
    ShortNameLength As Byte
    ShortName(23) As Byte
    FileName As Byte    '这里为了对齐，其实不用这个VB也是按4个字节对齐
    FileName1 As Integer    '(519) As Byte
End Type
Private Const FileBothDirectoryInformation = 3
Private Const SYNCHRONIZE = &H100000
Private Const FILE_ANY_ACCESS = 0
Private Const FILE_LIST_DIRECTORY = 1
Private Const FILE_DIRECTORY_FILE = 1
Private Const FILE_SYNCHRONOUS_IO_NONALERT = &H20
Private Const FILE_OPEN_FOR_BACKUP_INTENT = &H4000
Private Const OBJ_CASE_INSENSITIVE = &H40
Private Declare Function NtQueryDirectoryFile Lib "ntdll.dll" (ByVal FileHandle As Long, _
        ByVal hEvent As Long, _
        ByVal ApcRoutine As Long, _
        ByVal ApcContext As Long, _
        ByRef IoStatusBlock As Any, _
        ByRef FileInformation As Any, _
        ByVal Length As Long, _
        ByVal FileInformationClass As Long, _
        ByVal ReturnSingleEntry As Long, _
        ByRef FileName As Any, _
        ByVal RestartScan As Long _
        ) As Long
Private Declare Function NtClose Lib "ntdll.dll" (ByVal ObjectHandle As Long) As Long
Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function NtOpenFile Lib "ntdll.dll" (ByRef FileHandle As Long, _
        ByVal DesiredAccess As Long, _
        ByRef ObjectAttributes As OBJECT_ATTRIBUTES, _
        ByRef IoStatusBlock As IO_STATUS_BLOCK, _
        ByVal ShareAccess As Long, _
        ByVal OpenOptions As Long _
        ) As Long
Private Declare Sub RtlInitUnicodeString Lib "ntdll.dll" (DestinationString As Any, ByVal SourceString As Long)
Private Declare Sub ZeroMemory Lib "ntdll.dll" Alias "RtlZeroMemory" (dest As Any, ByVal numBytes As Long)
Private Function FindFirstFile(ByVal strDirectory As String, bytBuffer() As Byte) As Long
Dim strFolder As String
Dim obAttr As OBJECT_ATTRIBUTES
Dim objIoStatus As IO_STATUS_BLOCK
Dim ntStatus As Long
Dim hFind As Long
Dim strUnicode As UNICODE_STRING

    strFolder = "\??\"
    strFolder = strFolder & strDirectory
    RtlInitUnicodeString strUnicode, StrPtr(strFolder)    '初始化Unicode字符串
    obAttr.Length = LenB(obAttr)    '初始化OBJECT_ATTRIBUTES结构
    obAttr.Attributes = OBJ_CASE_INSENSITIVE
    obAttr.ObjectName = VarPtr(strUnicode)    '需要打开的文件夹路径
    obAttr.RootDirectory = 0
    obAttr.SecurityDescriptor = 0
    obAttr.SecurityQualityOfService = 0
    '获取文件夹句柄
    ntStatus = NtOpenFile(hFind, _
            FILE_LIST_DIRECTORY Or SYNCHRONIZE Or FILE_ANY_ACCESS, _
            obAttr, _
            objIoStatus, _
            3, _
            FILE_DIRECTORY_FILE Or FILE_SYNCHRONOUS_IO_NONALERT Or FILE_OPEN_FOR_BACKUP_INTENT)
    If ntStatus = 0 And hFind <> -1 Then
        '获取文件夹文件/目录信息,其实这个函数是Kernel32里的FindFirstFile的封装
        ntStatus = NtQueryDirectoryFile(hFind, _
                0, _
                0, _
                0, _
                objIoStatus, _
                bytBuffer(0), _
                UBound(bytBuffer), _
                FileBothDirectoryInformation, _
                1, _
                ByVal 0&, _
                0)
        If ntStatus = 0 Then    'Nt系列函数一般返回大于零表示成功，而NtQueryDirectoryFile返回零表示成功
            FindFirstFile = hFind
        Else
            NtClose hFind
        End If
    End If
End Function
Private Function FindNextFile(ByVal hFind As Long, bytBuffer() As Byte) As Boolean
Dim ntStatus As Long
Dim objIoStatus As IO_STATUS_BLOCK
    '这是Kernel32 FindNextFile的封装，只是是一次把所有文件/目录都取出来了
    ntStatus = NtQueryDirectoryFile(hFind, _
            0, _
            0, _
            0, _
            objIoStatus, _
            bytBuffer(0), _
            UBound(bytBuffer), _
            FileBothDirectoryInformation, _
            0, _
            ByVal 0&, _
            0)
    If ntStatus = 0 Then
        FindNextFile = True
    Else
        FindNextFile = False
    End If
End Function
Public Function EnumFile(ByVal strPath As String, ByVal LstOne As ListBox) As Integer
Dim pDir As FILE_BOTH_DIRECTORY_INFORMATION
Dim hFind As Long
Dim bytBuffer() As Byte
Dim bytName() As Byte
Dim strFileName As String * 520
Dim dwFileNameOffset As Long
Dim dwDirOffset As Long
    LstOne.Clear
    ReDim bytBuffer(LenB(pDir) + 260 * 2 - 3)    '定义一个缓存，返回数据用的
    If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
    '这里传字节数据方便处理些，本来打算直接使用FILE_BOTH_DIRECTORY_INFORMATION结构的但是发现VB在处理这方面比较差
    '尤其是在处理在遍历缓存中的所有文件和目录的时候更加麻烦了，因为FILE_BOTH_DIRECTORY_INFORMATION的长度其实是可
    '变的长度是根据LenB(FILE_BOTH_DIRECTORY_INFORMATION)+FileNameLength，这样直接用结构在VB中就无法定义了
    hFind = FindFirstFile(strPath, bytBuffer)    '获取第一个文件/目录对象
    CopyMemory pDir, bytBuffer(0), LenB(pDir)    '获取FILE_BOTH_DIRECTORY_INFORMATION结构，目的是获取FileNameLength和NextEntryOffset数据
    ReDim bytName(pDir.FileNameLength - 1)    '定义个和文件名一样大的缓存来存放文件名
    dwFileNameOffset = VarPtr(bytBuffer(&H5E))    '文件名偏移是&H5E
    CopyMemory bytName(0), ByVal dwFileNameOffset, pDir.FileNameLength
    strFileName = strPath & CStr(bytName)
    LstOne.AddItem strFileName
    Erase bytBuffer
    ReDim bytBuffer((LenB(pDir) + CLng(260 * 2 - 3)) * CLng(&H2000))
    If FindNextFile(hFind, bytBuffer) Then
        dwDirOffset = 0
        '遍历缓存
        Do While 1
            ZeroMemory pDir, LenB(pDir)    '把FILE_BOTH_DIRECTORY_INFORMATION置0
            CopyMemory pDir, ByVal VarPtr(bytBuffer(dwDirOffset)), LenB(pDir)    '得到FILE_BOTH_DIRECTORY_INFORMATION结构
            Erase bytName
            ReDim bytName(pDir.FileNameLength - 1)
            dwFileNameOffset = dwDirOffset + &H5E
            dwFileNameOffset = VarPtr(bytBuffer(dwFileNameOffset))
            CopyMemory bytName(0), ByVal dwFileNameOffset, pDir.FileNameLength
            strFileName = strPath & CStr(bytName)
            LstOne.AddItem strFileName    '这里如果是目录可以递归遍历目录下所有文件/目录,自己封装一个递归函数吧
            If pDir.NextEntryOffset = 0 Then Exit Do
            dwDirOffset = dwDirOffset + pDir.NextEntryOffset    '这里指向下一个FILE_BOTH_DIRECTORY_INFORMATION结构在内存中的位置
        Loop
    End If
    NtClose hFind
    EnumFile = LstOne.ListCount '文件总数
End Function

