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
   StartUpPosition =   3  '����ȱʡ
End
Attribute VB_Name = "frmDoku2Media"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'===================================================================================================
'����:����ٶȱ���ָ���ļ���
'�÷�:MsgBox EnumFile("C:\", List1)
'����:��C:\�µ������ļ���ʾ��List1��  ���������ļ���Ŀ
'���:NtQueryDirectoryFile��VB����DIR�Լ�FindFirstFile,FindNextFile��API���ں�ʵ�ֺ���.
'     �������Ÿ�3��ʹ��ǰ,�Ǿ�����װ�˵�,�Լ�ʵ����3���ķ�װ.
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
    FileName As Byte    '����Ϊ�˶��룬��ʵ�������VBҲ�ǰ�4���ֽڶ���
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
    RtlInitUnicodeString strUnicode, StrPtr(strFolder)    '��ʼ��Unicode�ַ���
    obAttr.Length = LenB(obAttr)    '��ʼ��OBJECT_ATTRIBUTES�ṹ
    obAttr.Attributes = OBJ_CASE_INSENSITIVE
    obAttr.ObjectName = VarPtr(strUnicode)    '��Ҫ�򿪵��ļ���·��
    obAttr.RootDirectory = 0
    obAttr.SecurityDescriptor = 0
    obAttr.SecurityQualityOfService = 0
    '��ȡ�ļ��о��
    ntStatus = NtOpenFile(hFind, _
            FILE_LIST_DIRECTORY Or SYNCHRONIZE Or FILE_ANY_ACCESS, _
            obAttr, _
            objIoStatus, _
            3, _
            FILE_DIRECTORY_FILE Or FILE_SYNCHRONOUS_IO_NONALERT Or FILE_OPEN_FOR_BACKUP_INTENT)
    If ntStatus = 0 And hFind <> -1 Then
        '��ȡ�ļ����ļ�/Ŀ¼��Ϣ,��ʵ���������Kernel32���FindFirstFile�ķ�װ
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
        If ntStatus = 0 Then    'Ntϵ�к���һ�㷵�ش������ʾ�ɹ�����NtQueryDirectoryFile�������ʾ�ɹ�
            FindFirstFile = hFind
        Else
            NtClose hFind
        End If
    End If
End Function
Private Function FindNextFile(ByVal hFind As Long, bytBuffer() As Byte) As Boolean
Dim ntStatus As Long
Dim objIoStatus As IO_STATUS_BLOCK
    '����Kernel32 FindNextFile�ķ�װ��ֻ����һ�ΰ������ļ�/Ŀ¼��ȡ������
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
    ReDim bytBuffer(LenB(pDir) + 260 * 2 - 3)    '����һ�����棬���������õ�
    If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
    '���ﴫ�ֽ����ݷ��㴦��Щ����������ֱ��ʹ��FILE_BOTH_DIRECTORY_INFORMATION�ṹ�ĵ��Ƿ���VB�ڴ����ⷽ��Ƚϲ�
    '�������ڴ����ڱ��������е������ļ���Ŀ¼��ʱ������鷳�ˣ���ΪFILE_BOTH_DIRECTORY_INFORMATION�ĳ�����ʵ�ǿ�
    '��ĳ����Ǹ���LenB(FILE_BOTH_DIRECTORY_INFORMATION)+FileNameLength������ֱ���ýṹ��VB�о��޷�������
    hFind = FindFirstFile(strPath, bytBuffer)    '��ȡ��һ���ļ�/Ŀ¼����
    CopyMemory pDir, bytBuffer(0), LenB(pDir)    '��ȡFILE_BOTH_DIRECTORY_INFORMATION�ṹ��Ŀ���ǻ�ȡFileNameLength��NextEntryOffset����
    ReDim bytName(pDir.FileNameLength - 1)    '��������ļ���һ����Ļ���������ļ���
    dwFileNameOffset = VarPtr(bytBuffer(&H5E))    '�ļ���ƫ����&H5E
    CopyMemory bytName(0), ByVal dwFileNameOffset, pDir.FileNameLength
    strFileName = strPath & CStr(bytName)
    LstOne.AddItem strFileName
    Erase bytBuffer
    ReDim bytBuffer((LenB(pDir) + CLng(260 * 2 - 3)) * CLng(&H2000))
    If FindNextFile(hFind, bytBuffer) Then
        dwDirOffset = 0
        '��������
        Do While 1
            ZeroMemory pDir, LenB(pDir)    '��FILE_BOTH_DIRECTORY_INFORMATION��0
            CopyMemory pDir, ByVal VarPtr(bytBuffer(dwDirOffset)), LenB(pDir)    '�õ�FILE_BOTH_DIRECTORY_INFORMATION�ṹ
            Erase bytName
            ReDim bytName(pDir.FileNameLength - 1)
            dwFileNameOffset = dwDirOffset + &H5E
            dwFileNameOffset = VarPtr(bytBuffer(dwFileNameOffset))
            CopyMemory bytName(0), ByVal dwFileNameOffset, pDir.FileNameLength
            strFileName = strPath & CStr(bytName)
            LstOne.AddItem strFileName    '���������Ŀ¼���Եݹ����Ŀ¼�������ļ�/Ŀ¼,�Լ���װһ���ݹ麯����
            If pDir.NextEntryOffset = 0 Then Exit Do
            dwDirOffset = dwDirOffset + pDir.NextEntryOffset    '����ָ����һ��FILE_BOTH_DIRECTORY_INFORMATION�ṹ���ڴ��е�λ��
        Loop
    End If
    NtClose hFind
    EnumFile = LstOne.ListCount '�ļ�����
End Function

