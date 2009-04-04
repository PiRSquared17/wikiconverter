VERSION 5.00
Begin VB.Form frmGb2utf 
   Caption         =   "GB/BIG/UTF-8 �ļ�����ת��"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7875
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5895
   ScaleWidth      =   7875
   StartUpPosition =   2  '��Ļ����
   Begin VB.CheckBox chkBackup 
      Caption         =   "�����ļ�����"
      Height          =   360
      Left            =   84
      TabIndex        =   11
      Top             =   5055
      Value           =   1  'Checked
      Width           =   1425
   End
   Begin VB.CheckBox chkText 
      Caption         =   "�滻����..."
      Height          =   360
      Left            =   84
      TabIndex        =   12
      Top             =   5400
      Width           =   1425
   End
   Begin VB.Frame freFrmCode 
      Caption         =   "ԭ�ļ�����"
      Height          =   705
      Left            =   1649
      TabIndex        =   22
      Top             =   4995
      Width           =   2500
      Begin VB.OptionButton optFrmGb 
         Caption         =   "GB"
         Height          =   390
         Left            =   982
         TabIndex        =   14
         Top             =   255
         Value           =   -1  'True
         Width           =   600
      End
      Begin VB.OptionButton optFrmUTF8 
         Caption         =   "UTF-8"
         Height          =   390
         Left            =   105
         TabIndex        =   13
         Top             =   255
         Width           =   810
      End
      Begin VB.OptionButton optFrmBig5 
         Caption         =   "BIG5"
         Height          =   390
         Left            =   1650
         TabIndex        =   15
         Top             =   255
         Width           =   720
      End
   End
   Begin VB.Frame freCode 
      Caption         =   "ת�������"
      Height          =   705
      Left            =   4305
      TabIndex        =   21
      Top             =   4995
      Width           =   3465
      Begin VB.CheckBox chkU8BOM 
         Caption         =   "��BOM"
         Height          =   285
         Left            =   2520
         TabIndex        =   23
         Top             =   308
         Value           =   1  'Checked
         Width           =   825
      End
      Begin VB.OptionButton optGb 
         Caption         =   "GB"
         Height          =   390
         Left            =   975
         TabIndex        =   17
         Top             =   255
         Width           =   600
      End
      Begin VB.OptionButton optUtf8 
         Caption         =   "UTF-8"
         Height          =   390
         Left            =   90
         TabIndex        =   16
         Top             =   255
         Value           =   -1  'True
         Width           =   810
      End
      Begin VB.OptionButton optBig5 
         Caption         =   "BIG5"
         Height          =   390
         Left            =   1650
         TabIndex        =   18
         Top             =   255
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdRemoveUnsel 
      Caption         =   "�����Ƴ�(&U)"
      Enabled         =   0   'False
      Height          =   360
      Left            =   6396
      TabIndex        =   7
      Top             =   2812
      Width           =   1400
   End
   Begin VB.Frame cmdConvertInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   1684
      TabIndex        =   19
      Top             =   2212
      Visible         =   0   'False
      Width           =   3000
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ڴ������Ժ�..."
         Height          =   180
         Left            =   660
         TabIndex        =   20
         Top             =   255
         Width           =   1710
      End
   End
   Begin VB.CommandButton cmdStopConvert 
      Caption         =   "ֹͣ����(&C)"
      Enabled         =   0   'False
      Height          =   360
      Left            =   6396
      TabIndex        =   3
      Top             =   825
      Width           =   1400
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "����wikiת��"
      Height          =   360
      Left            =   6396
      TabIndex        =   9
      Top             =   3825
      Width           =   1400
   End
   Begin VB.CommandButton cmdRemoveAll 
      Caption         =   "ȫ���Ƴ�(&A)"
      Enabled         =   0   'False
      Height          =   360
      Left            =   6396
      TabIndex        =   8
      Top             =   3225
      Width           =   1400
   End
   Begin VB.CommandButton cmdRemoveFile 
      Caption         =   "�Ƴ�ѡ��(&R)"
      Enabled         =   0   'False
      Height          =   360
      Left            =   6396
      TabIndex        =   6
      Top             =   2385
      Width           =   1400
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "��    ��(&X)"
      Height          =   360
      Left            =   6396
      TabIndex        =   10
      Top             =   4252
      Width           =   1400
   End
   Begin VB.CommandButton cmdStartConvert 
      Caption         =   "��ʼ����(&S)"
      Enabled         =   0   'False
      Height          =   360
      Left            =   6396
      TabIndex        =   2
      Top             =   412
      Width           =   1400
   End
   Begin VB.CommandButton cmdAddDir 
      Caption         =   "���Ŀ¼..."
      Height          =   360
      Left            =   6396
      TabIndex        =   5
      Top             =   1815
      Width           =   1400
   End
   Begin VB.CommandButton cmdAddFile 
      Caption         =   "����ļ�..."
      Height          =   360
      Left            =   6396
      TabIndex        =   4
      Top             =   1395
      Width           =   1400
   End
   Begin VB.ListBox lstFiles 
      Height          =   4200
      Left            =   84
      MultiSelect     =   2  'Extended
      TabIndex        =   1
      ToolTipText     =   "˫����س��鿴�ļ�����"
      Top             =   412
      Width           =   6200
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   7947
      Y1              =   4785
      Y2              =   4785
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   7927
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Label lblFiles 
      AutoSize        =   -1  'True
      Caption         =   "�ļ��б�(ע�ⱸ��)��"
      Height          =   180
      Left            =   75
      TabIndex        =   0
      Top             =   135
      Width           =   1800
   End
End
Attribute VB_Name = "frmGb2utf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'---------------------------------------------------
'���弶���ñ���
'---------------------------------------------------
Private pSelPath As String, pSelExt As String, logFile As String
Private BreakConvert As Boolean

'---------------------------------------------------
'; �����Զ����ű���
'---------------------------------------------------
Private arrSize() As ControlSize

Private Type ControlSize
    name As String          '�ؼ�����
    X As Single             '��߾����
    Y As Single             '�ұ߾����
    cx As Single            '��ȱ���
    cy As Single            '���ȱ���
End Type

'---------------------------------------------------
'�Ƿ��м��� & ��궯��
'--------------------API��������--------------------
Private Declare Function GetInputState Lib "user32" () As Long

'---------------------------------------------------
'�Զ�����ҳ
'--------------------API��������--------------------
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

';**********************************
'; ����Ϊ�¼�����
';**********************************
Private Sub chkText_Click()
    Dim sCode As String, sPage As String
    
    '�������ݣ����ļ�ͷ���һЩ asp ����
    If FileTextAction(0).TextAct = -1 Then
        If optUtf8.Value Then
            sCode = "UTF-8"
            sPage = "65001"
        End If
        If optGb.Value Then
            sCode = "GB2312"
            sPage = "936"
        End If
        If optBig5.Value Then
            sCode = "BIG5"
            sPage = "950"
        End If
        
        FileTextAction(0).TextAct = 1
        FileTextAction(0).OldText = "<�����׷�������ݣ����ﲻ��Ҫ����>"
        FileTextAction(0).NewText = "<%@ CODEPAGE=" & sPage & " %>" & vbCrLf & "<%Response.Charset=""" & sCode & """%>" & vbCrLf & vbCrLf _
                    & "<!-- (ע��)���ļ��� GB2UTF8 1.21 ������ " & Now() & "��������ߣ�����(fxy_2002@163.com) pc-soft.cn -->" & vbCrLf
    End If

    '��ʾ�����崰��
    If chkText.Value Then frmText.Show vbModal
End Sub

Private Sub cmdRemoveUnsel_Click()
    RemoveFiles True
End Sub

Private Sub Form_Load()
    
    ReDim arrSize(0) As ControlSize
    
    SaveSize lblFiles
    SaveSize lstFiles
    SaveSize cmdConvertInfo
    SaveSize lblInfo
    SaveSize cmdStartConvert
    SaveSize cmdStopConvert
    SaveSize cmdAddFile
    SaveSize cmdAddDir
    SaveSize cmdRemoveFile
    SaveSize cmdRemoveUnsel
    SaveSize cmdRemoveAll
    SaveSize cmdAbout
    SaveSize Cmdexit

    SaveSize chkBackup
    SaveSize chkText
    SaveSize freFrmCode
    SaveSize freCode


    '��־�ļ�
    logFile = Replace(App.Path & "\GB2UTF8.log", "\\", "\")
    
    '�ϴδ���״̬���������󻯣��ָ���������ʾĬ�ϴ�С
    If GetSetting("GB2UTF8", "Parameter", "Window", 0) = 2 Then Me.WindowState = 2

End Sub



Private Sub Form_Resize()
    Select Case Me.WindowState
        Case 1
            Exit Sub
        Case 0
            If Me.Width < 8000 Or Me.Height < 6400 Then
                Me.Move Me.Left, Me.Top, 8000, 6400
            Else
                Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
            End If
        Case Else
            '
    End Select
    
    Resize lstFiles
    Resize lblFiles
    Resize cmdConvertInfo
    Resize lblInfo
    Resize cmdStartConvert
    Resize cmdStopConvert
    Resize cmdAddFile
    Resize cmdAddDir
    Resize cmdRemoveFile
    Resize cmdRemoveUnsel
    Resize cmdRemoveAll
    Resize cmdAbout
    Resize Cmdexit
    
    Resize chkBackup
    Resize chkText
    Resize freFrmCode
    Resize freCode

End Sub

Private Sub cmdAbout_Click()
    Load frmMain
    frmMain.Show
    'Unload frmGb2utf
End Sub

Private Sub cmdAddDir_Click()
    Dim gDir As String, sExt As String
    Dim aFiles() As String, i As Long
    
    gDir = BrowseForFolder("ѡ��Ŀ¼...", pSelPath, False, Me.hWnd)
    If Len(gDir) = 0 Then Exit Sub
    pSelPath = gDir
    
    If Len(pSelExt) = 0 Then pSelExt = "*.htm|*.html|*.vbs|*.js|*.asp|*.txt"
    sExt = InputBox("�ļ����ͣ�" & vbCrLf & "(��ͬ������ ��|������;����,���ָ�)", "��������", Replace(Replace(pSelExt, ";", "|"), ",", "|"))
    If Len(sExt) = 0 Then Exit Sub
    pSelExt = sExt
    
    lblInfo.Caption = "�����ļ������Ժ�..."
    cmdConvertInfo.Visible = True
    DoEvents
    
    aFiles = SearchFiles(gDir, sExt, True)
    If Len(aFiles(0)) > 0 Then
        For i = 0 To UBound(aFiles)
            'If Not ExistInList(lstFiles, aFiles(i)) Then
                lstFiles.AddItem aFiles(i)
            'End If
        Next
        cmdStartConvert.Enabled = True
        cmdRemoveFile.Enabled = True
        cmdRemoveUnsel.Enabled = True
        cmdRemoveAll.Enabled = True
    End If

    cmdConvertInfo.Visible = False
End Sub

Private Sub cmdAddFile_Click()
    Dim gFiles As String
    Dim aFiles() As String, i As Long
    
    gFiles = GetFileNames(ShowOpenFile, "����ļ�...", , "��ҳ�ļ�(*.htm; *.html)|*.htm;*.html|�ű��ļ�(*.vbs; *.js)|*.vbs;*.js|ASP�ļ�(*.asp)|*.asp|�ı��ļ�(*.txt)|*.txt|�����ļ�(*.*)|*.*", True, Me.hWnd)
    
    If Len(gFiles) > 0 Then
        aFiles = Split(gFiles, "|")
        For i = 0 To UBound(aFiles)
            'If Not ExistInList(lstFiles, aFiles(i)) Then
                lstFiles.AddItem aFiles(i)
            'End If
        Next
        cmdStartConvert.Enabled = True
        cmdRemoveFile.Enabled = True
        cmdRemoveUnsel.Enabled = True
        cmdRemoveAll.Enabled = True
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdRemoveAll_Click()
    lstFiles.Clear
    
    cmdStartConvert.Enabled = False
    cmdRemoveFile.Enabled = False
    cmdRemoveUnsel.Enabled = False
    cmdRemoveAll.Enabled = False
End Sub

Private Sub cmdRemoveFile_Click()
    If lstFiles.SelCount > 0 Then RemoveFiles
End Sub

Private Sub cmdStartConvert_Click()
    Dim i As Long, f As Long
    Dim p As Long, s As Long, psi As Long
    Dim flag As Boolean, bak As Boolean
    Dim frmCode As TextCode, toCode As TextCode
    Dim frmcName As String, tocName As String
    Dim e As String
    
    '��Դ�ļ������ʽ
    If optFrmUTF8.Value Then
        frmCode = UTF8_Code
        frmcName = "UTF-8"
    End If
    If optFrmGb.Value Then
        frmCode = GB2312_Code
        frmcName = "GB"
    End If
    If optFrmBig5.Value Then
        frmCode = BIG5_Code
        frmcName = "BIG5"
    End If
    
    'Ŀ�����
    If optUtf8.Value Then
        toCode = UTF8_Code
        tocName = "UTF-8"
    End If
    If optGb.Value Then
        toCode = GB2312_Code
        tocName = "GB"
    End If
    If optBig5.Value Then
        toCode = BIG5_Code
        tocName = "BIG5"
    End If

    If frmCode = toCode And chkText.Value = 0 Then
        '���벻���ı䣬�����ļ����ݲ����滻
        
        MsgBox "�����ʽ���ļ����ݾ������ı䣬����Ҫ����", vbOKOnly + vbCritical, "��ʾ"
        Exit Sub
        
    End If

    If chkBackup.Value Then
        bak = True
    Else
        '���������ݣ���ʾ������Ϣ
        If MsgBoxA("������̽���дԴ�ļ����Ҳ�����ʾ�������ñ��ݴ�ʩ����Ȼ��������ʱ��ֹ������̣����Ѵ�����ļ��������ٻ�ԭ��ȷ��Ҫ������", "OK��֪����,���ٿ���", "��ʼ", vbOKCancel + vbQuestion + vbDefaultButton2, Me.hWnd) = vbCancel Then Exit Sub
    End If
    
    unSelectedAll   'ȡ������ѡ�񣬱���ѡ������ٵ�ǰ�����ļ�
    
    'ʹ���������ֹ��ť��������ؼ����ɲ���
    lstFiles.Enabled = False
    cmdStartConvert.Enabled = False
    chkBackup.Enabled = False
    chkText.Enabled = False

    cmdStopConvert.Enabled = True
    cmdAddFile.Enabled = False
    cmdAddDir.Enabled = False
    cmdRemoveFile.Enabled = False
    cmdRemoveUnsel.Enabled = False
    cmdRemoveAll.Enabled = False
    'cmdAbout.Enabled = False
    Cmdexit.Enabled = False
    
    lblInfo.Caption = "���ڴ������Ժ�..."
    cmdConvertInfo.Visible = True
    DoEvents
    
    '��ʼ����
    e = "<GB2UTF8 ������־>" & vbCrLf & "��ʼʱ�䣺" & Now() & vbCrLf
    e = e & "ִ�в�����" & frmcName & " --> " & tocName & vbCrLf & vbCrLf
    
    BreakConvert = False
    
    s = lstFiles.ListCount
    p = s \ 100 + 1
    For i = 0 To s - 1
        If (i + 1) Mod p = 0 Then       '��������ﵽ 1% �Ĳ�����������ʾ���ȣ����� doevents ���Ĵ����ӿ촦��
            lblInfo.Caption = "���ڴ�������� " & CInt((i + 1) * 100 \ s) & " %"
            
            lstFiles.Selected(psi) = False  'ȡ���ϴ�ѡ�����Ŀ
            psi = i
            
            lstFiles.Selected(i) = True     'ѡ�����ڴ�����ļ������׹������ʾ�û�
            lstFiles.TopIndex = i
            
            flag = True
        End If
        
        If GetInputState() Or flag Then
            DoEvents
            flag = False
        End If
        
        If BreakConvert Then    '�û�����ֹͣ������ ESC ��
            If MsgBox("ȷ��Ҫ��ֹ��ǰ�Ĵ��������", vbOKCancel + vbQuestion + vbDefaultButton2, "�ж�") = vbOK Then
                e = e & "*** ������̱��û���ֹ�� ***" & vbCrLf & lstFiles.List(i) & vbCrLf & vbCrLf & "��ֹʱ�䣺" & Now()
                'lstFiles.Selected(i) = True
                Exit For
            Else
                BreakConvert = False
            End If
        End If
        
        '�����ļ�
        If Not Save2File(lstFiles.List(i), frmCode, toCode, bak, chkText.Value, chkU8BOM.Value) Then
            e = e & "����ʧ�ܣ�" & lstFiles.List(i) & vbCrLf   '������Ϣ
            f = f + 1       'ʧ�ܼ���
        End If
    Next
    
    lstFiles.Enabled = True
    cmdStartConvert.Enabled = True
    chkBackup.Enabled = True
    chkText.Enabled = True
    
    cmdStopConvert.Enabled = False
    cmdConvertInfo.Visible = False
    cmdAddFile.Enabled = True
    cmdAddDir.Enabled = True
    cmdRemoveFile.Enabled = True
    cmdRemoveUnsel.Enabled = True
    cmdRemoveAll.Enabled = True
    'cmdAbout.Enabled = True
    Cmdexit.Enabled = True
    
    If BreakConvert Then
        ErrLog e
        MsgBox "������ֹ��", vbOKOnly + vbExclamation, "�ж�"
    Else
        e = e & vbCrLf & "���ʱ�䣺" & Now()
        ErrLog e
        If f = 0 Then
            MsgBox "�ļ�������ɡ�", vbOKOnly + vbInformation, "���"
        Else
            MsgBox "������ɣ��� " & CStr(f) & " ���ļ�����ʧ�ܡ�������Ϣ��鿴�ļ���" & logFile, vbOKOnly + vbExclamation, "���"
            ViewFile logFile    '�Զ���ʾ��־�ļ�
        End If
    End If
End Sub

Private Sub cmdStopConvert_Click()
    BreakConvert = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then BreakConvert = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "GB2UTF8", "Parameter", "Window", Me.WindowState
    End
End Sub
Private Sub lstFiles_DblClick()
    ViewFile lstFiles.List(lstFiles.ListIndex)
End Sub

Private Sub lstFiles_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If lstFiles.ListIndex > -1 Then lstFiles_DblClick
    End If
End Sub

';**********************************
'; ����Ϊ�Զ��庯��
';**********************************
Private Sub Resize(obj As Object)
'�����ڿؼ��Զ�����

    Dim i As Integer
    Dim j As Integer
    j = UBound(arrSize())
    For i = 1 To j
        If arrSize(i).name = obj.name Then
            Exit For
        End If
    Next i
    If i > j Then
        Exit Sub
    Else
        obj.Left = arrSize(i).X * Me.ScaleWidth
        obj.Top = arrSize(i).Y * Me.ScaleHeight
        obj.Width = arrSize(i).cx * Me.ScaleWidth
        obj.Height = arrSize(i).cy * Me.ScaleHeight
    End If
    
End Sub

Private Sub SaveSize(obj As Object)
'���洰��λ�ñ���

    Dim count As Integer
    count = UBound(arrSize)
    ReDim Preserve arrSize(count + 1)
    arrSize(count + 1).name = obj.name
    arrSize(count + 1).X = obj.Left / Me.ScaleWidth
    arrSize(count + 1).Y = obj.Top / Me.ScaleHeight
    arrSize(count + 1).cx = obj.Width / Me.ScaleWidth
    arrSize(count + 1).cy = obj.Height / Me.ScaleHeight
End Sub

Private Sub RemoveFiles(Optional RemoveUnsel As Boolean)
'�Ƴ��б���Ŀ��
'Removeunsel ��ʾ�Ƴ�δѡ�񣬷����Ƴ�ѡ��

    Dim i As Long
    
    For i = lstFiles.ListCount - 1 To 0 Step -1
        If lstFiles.Selected(i) = Not RemoveUnsel Then
            lstFiles.RemoveItem i
        End If
    Next
    lstFiles.SetFocus
    
    If lstFiles.ListCount < 1 Then
        cmdStartConvert.Enabled = False
        cmdRemoveFile.Enabled = False
        cmdRemoveUnsel.Enabled = False
        cmdRemoveAll.Enabled = False
    End If
    
End Sub

Private Sub unSelectedAll()
    '��ѡ��������Ŀ
    Dim i As Long
    For i = 0 To lstFiles.ListCount - 1
        lstFiles.Selected(i) = False
    Next
End Sub

Private Sub ErrLog(ByVal errStr As String)
'����������־�ļ�
    Dim fh As Integer
    
    fh = FreeFile()
    Open logFile For Output As #fh
    Print #fh, errStr
    Close #fh
End Sub

Private Sub ViewFile(ByVal fName As String)
'�ü��±���ʾ�ļ�
On Error GoTo aErr

    If Len(fName) > 0 And Len(Dir(fName)) > 0 Then
        Shell "notepad.exe " & Chr(34) & fName & Chr(34), vbNormalFocus
    End If
    Exit Sub
    
aErr:
    MsgBox Err.Description, vbCritical + vbOKOnly, "����"
End Sub

Private Sub ToWebHome(ByVal sTarget As String)
'�Զ�����ҳ
    ShellExecute &H0, "open", sTarget, vbNullString, vbNullString, 1
End Sub

Private Sub optBig5_Click()
    chkU8BOM.Enabled = optUtf8.Value
End Sub

Private Sub optGb_Click()
    chkU8BOM.Enabled = optUtf8.Value
End Sub

'Private Function ExistInList(ByRef oList As ListBox, ByVal fName As String) As Boolean
''�ļ����Ƿ�����б��У��ж��Ƿ��ظ���ӣ�----��ͣ�˺������ٶ�Ӱ��̫��
'    Dim i As Long
'    If oList.ListCount > 0 Then
'        For i = 0 To oList.ListCount - 1
'            If LCase(oList.List(i)) = LCase(fName) Then
'                ExistInList = True
'                Exit For
'            End If
'        Next
'    End If
'End Function

Private Sub optUtf8_Click()
    chkU8BOM.Enabled = optUtf8.Value
End Sub
