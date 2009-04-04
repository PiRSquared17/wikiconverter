VERSION 5.00
Begin VB.Form frmGb2utf 
   Caption         =   "GB/BIG/UTF-8 文件编码转换"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7875
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5895
   ScaleWidth      =   7875
   StartUpPosition =   2  '屏幕中心
   Begin VB.CheckBox chkBackup 
      Caption         =   "保留文件备份"
      Height          =   360
      Left            =   84
      TabIndex        =   11
      Top             =   5055
      Value           =   1  'Checked
      Width           =   1425
   End
   Begin VB.CheckBox chkText 
      Caption         =   "替换内容..."
      Height          =   360
      Left            =   84
      TabIndex        =   12
      Top             =   5400
      Width           =   1425
   End
   Begin VB.Frame freFrmCode 
      Caption         =   "原文件编码"
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
      Caption         =   "转换后编码"
      Height          =   705
      Left            =   4305
      TabIndex        =   21
      Top             =   4995
      Width           =   3465
      Begin VB.CheckBox chkU8BOM 
         Caption         =   "带BOM"
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
      Caption         =   "反向移除(&U)"
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
         Caption         =   "正在处理，请稍侯..."
         Height          =   180
         Left            =   660
         TabIndex        =   20
         Top             =   255
         Width           =   1710
      End
   End
   Begin VB.CommandButton cmdStopConvert 
      Caption         =   "停止处理(&C)"
      Enabled         =   0   'False
      Height          =   360
      Left            =   6396
      TabIndex        =   3
      Top             =   825
      Width           =   1400
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "返回wiki转换"
      Height          =   360
      Left            =   6396
      TabIndex        =   9
      Top             =   3825
      Width           =   1400
   End
   Begin VB.CommandButton cmdRemoveAll 
      Caption         =   "全部移除(&A)"
      Enabled         =   0   'False
      Height          =   360
      Left            =   6396
      TabIndex        =   8
      Top             =   3225
      Width           =   1400
   End
   Begin VB.CommandButton cmdRemoveFile 
      Caption         =   "移除选择(&R)"
      Enabled         =   0   'False
      Height          =   360
      Left            =   6396
      TabIndex        =   6
      Top             =   2385
      Width           =   1400
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "退    出(&X)"
      Height          =   360
      Left            =   6396
      TabIndex        =   10
      Top             =   4252
      Width           =   1400
   End
   Begin VB.CommandButton cmdStartConvert 
      Caption         =   "开始处理(&S)"
      Enabled         =   0   'False
      Height          =   360
      Left            =   6396
      TabIndex        =   2
      Top             =   412
      Width           =   1400
   End
   Begin VB.CommandButton cmdAddDir 
      Caption         =   "添加目录..."
      Height          =   360
      Left            =   6396
      TabIndex        =   5
      Top             =   1815
      Width           =   1400
   End
   Begin VB.CommandButton cmdAddFile 
      Caption         =   "添加文件..."
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
      ToolTipText     =   "双击或回车查看文件内容"
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
      Caption         =   "文件列表(注意备份)："
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
'窗体级公用变量
'---------------------------------------------------
Private pSelPath As String, pSelExt As String, logFile As String
Private BreakConvert As Boolean

'---------------------------------------------------
'; 窗体自动缩放变量
'---------------------------------------------------
Private arrSize() As ControlSize

Private Type ControlSize
    name As String          '控件名称
    X As Single             '左边距比例
    Y As Single             '右边距比例
    cx As Single            '宽度比例
    cy As Single            '长度比例
End Type

'---------------------------------------------------
'是否有键盘 & 鼠标动作
'--------------------API声明部分--------------------
Private Declare Function GetInputState Lib "user32" () As Long

'---------------------------------------------------
'自动打开网页
'--------------------API声明部分--------------------
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

';**********************************
'; 以下为事件代码
';**********************************
Private Sub chkText_Click()
    Dim sCode As String, sPage As String
    
    '样板数据，在文件头添加一些 asp 代码
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
        FileTextAction(0).OldText = "<插入或追加新内容，这里不需要输入>"
        FileTextAction(0).NewText = "<%@ CODEPAGE=" & sPage & " %>" & vbCrLf & "<%Response.Charset=""" & sCode & """%>" & vbCrLf & vbCrLf _
                    & "<!-- (注释)此文件由 GB2UTF8 1.21 生成于 " & Now() & "，软件作者：阿勇(fxy_2002@163.com) pc-soft.cn -->" & vbCrLf
    End If

    '显示规则定义窗口
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


    '日志文件
    logFile = Replace(App.Path & "\GB2UTF8.log", "\\", "\")
    
    '上次窗口状态，如果是最大化，恢复；否则，显示默认大小
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
    
    gDir = BrowseForFolder("选择目录...", pSelPath, False, Me.hWnd)
    If Len(gDir) = 0 Then Exit Sub
    pSelPath = gDir
    
    If Len(pSelExt) = 0 Then pSelExt = "*.htm|*.html|*.vbs|*.js|*.asp|*.txt"
    sExt = InputBox("文件类型：" & vbCrLf & "(不同类型以 “|”、“;”或“,”分隔)", "过滤条件", Replace(Replace(pSelExt, ";", "|"), ",", "|"))
    If Len(sExt) = 0 Then Exit Sub
    pSelExt = sExt
    
    lblInfo.Caption = "查找文件，请稍侯..."
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
    
    gFiles = GetFileNames(ShowOpenFile, "添加文件...", , "网页文件(*.htm; *.html)|*.htm;*.html|脚本文件(*.vbs; *.js)|*.vbs;*.js|ASP文件(*.asp)|*.asp|文本文件(*.txt)|*.txt|所有文件(*.*)|*.*", True, Me.hWnd)
    
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
    
    '来源文件编码格式
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
    
    '目标编码
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
        '编码不作改变，并且文件内容不作替换
        
        MsgBox "编码格式和文件内容均不做改变，不需要处理。", vbOKOnly + vbCritical, "提示"
        Exit Sub
        
    End If

    If chkBackup.Value Then
        bak = True
    Else
        '不保留备份，显示警告信息
        If MsgBoxA("处理过程将覆写源文件并且不再提示，请做好备份措施。虽然您可以随时终止处理过程，但已处理的文件将不可再还原。确定要继续吗？", "OK，知道了,我再看看", "开始", vbOKCancel + vbQuestion + vbDefaultButton2, Me.hWnd) = vbCancel Then Exit Sub
    End If
    
    unSelectedAll   '取消所有选择，便于选择光标跟踪当前处理文件
    
    '使界面除了终止按钮外的其它控件不可操作
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
    
    lblInfo.Caption = "正在处理，请稍侯..."
    cmdConvertInfo.Visible = True
    DoEvents
    
    '开始处理
    e = "<GB2UTF8 运行日志>" & vbCrLf & "开始时间：" & Now() & vbCrLf
    e = e & "执行操作：" & frmcName & " --> " & tocName & vbCrLf & vbCrLf
    
    BreakConvert = False
    
    s = lstFiles.ListCount
    p = s \ 100 + 1
    For i = 0 To s - 1
        If (i + 1) Mod p = 0 Then       '处理比例达到 1% 的不整数倍才显示进度，减少 doevents 语句的次数加快处理
            lblInfo.Caption = "正在处理，已完成 " & CInt((i + 1) * 100 \ s) & " %"
            
            lstFiles.Selected(psi) = False  '取消上次选择的项目
            psi = i
            
            lstFiles.Selected(i) = True     '选择正在处理的文件，反白光标以提示用户
            lstFiles.TopIndex = i
            
            flag = True
        End If
        
        If GetInputState() Or flag Then
            DoEvents
            flag = False
        End If
        
        If BreakConvert Then    '用户按下停止按按或 ESC 键
            If MsgBox("确定要终止当前的处理过程吗？", vbOKCancel + vbQuestion + vbDefaultButton2, "中断") = vbOK Then
                e = e & "*** 处理过程被用户终止于 ***" & vbCrLf & lstFiles.List(i) & vbCrLf & vbCrLf & "终止时间：" & Now()
                'lstFiles.Selected(i) = True
                Exit For
            Else
                BreakConvert = False
            End If
        End If
        
        '处理文件
        If Not Save2File(lstFiles.List(i), frmCode, toCode, bak, chkText.Value, chkU8BOM.Value) Then
            e = e & "处理失败：" & lstFiles.List(i) & vbCrLf   '错误信息
            f = f + 1       '失败计数
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
        MsgBox "处理被终止。", vbOKOnly + vbExclamation, "中断"
    Else
        e = e & vbCrLf & "完成时间：" & Now()
        ErrLog e
        If f = 0 Then
            MsgBox "文件处理完成。", vbOKOnly + vbInformation, "完成"
        Else
            MsgBox "处理完成，有 " & CStr(f) & " 个文件处理失败。更多信息请查看文件：" & logFile, vbOKOnly + vbExclamation, "完成"
            ViewFile logFile    '自动显示日志文件
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
'; 以下为自定义函数
';**********************************
Private Sub Resize(obj As Object)
'窗件内控件自动缩放

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
'保存窗体位置比例

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
'移除列表项目：
'Removeunsel 表示移除未选择，否则移除选择

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
    '不选择所有项目
    Dim i As Long
    For i = 0 To lstFiles.ListCount - 1
        lstFiles.Selected(i) = False
    Next
End Sub

Private Sub ErrLog(ByVal errStr As String)
'建立错误日志文件
    Dim fh As Integer
    
    fh = FreeFile()
    Open logFile For Output As #fh
    Print #fh, errStr
    Close #fh
End Sub

Private Sub ViewFile(ByVal fName As String)
'用记事本显示文件
On Error GoTo aErr

    If Len(fName) > 0 And Len(Dir(fName)) > 0 Then
        Shell "notepad.exe " & Chr(34) & fName & Chr(34), vbNormalFocus
    End If
    Exit Sub
    
aErr:
    MsgBox Err.Description, vbCritical + vbOKOnly, "错误"
End Sub

Private Sub ToWebHome(ByVal sTarget As String)
'自动打开网页
    ShellExecute &H0, "open", sTarget, vbNullString, vbNullString, 1
End Sub

Private Sub optBig5_Click()
    chkU8BOM.Enabled = optUtf8.Value
End Sub

Private Sub optGb_Click()
    chkU8BOM.Enabled = optUtf8.Value
End Sub

'Private Function ExistInList(ByRef oList As ListBox, ByVal fName As String) As Boolean
''文件名是否存在列表中（判断是否重复添加）----暂停此函数，速度影响太大
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
