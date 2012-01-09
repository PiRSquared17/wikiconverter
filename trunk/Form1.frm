VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Wiki converter"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   7095
   StartUpPosition =   1  '所有者中心
   Begin TabDlg.SSTab SSTab1 
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   12091
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Openwiki转换成Dokuwiki"
      TabPicture(0)   =   "Form1.frx":1CFA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(1)=   "Label7"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Dokuwiki转换成Mediawiki"
      TabPicture(1)   =   "Form1.frx":1D16
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame2 
         Caption         =   "Dokuwiki的TXT页面转换成Mediawiki的import.xml"
         Height          =   5775
         Left            =   360
         TabIndex        =   20
         Top             =   600
         Width           =   6015
         Begin VB.CommandButton btnDEUtf8 
            Caption         =   "文件名批量解码UTF8"
            Height          =   495
            Left            =   4560
            TabIndex        =   29
            Top             =   3720
            Width           =   1215
         End
         Begin VB.CommandButton btnConvert 
            Caption         =   "文件批量导入XML"
            Height          =   615
            Left            =   4560
            TabIndex        =   28
            Top             =   4440
            Width           =   1215
         End
         Begin VB.ListBox List2 
            Height          =   1320
            Left            =   600
            TabIndex        =   27
            Top             =   3720
            Width           =   3615
         End
         Begin VB.TextBox txtCurrentDir 
            Height          =   615
            Left            =   1080
            MultiLine       =   -1  'True
            TabIndex        =   25
            Top             =   2760
            Width           =   4215
         End
         Begin VB.CommandButton cmdup 
            Caption         =   "返回上级"
            Height          =   315
            Left            =   4560
            TabIndex        =   24
            Top             =   360
            Width           =   1275
         End
         Begin VB.ComboBox cmbDrives 
            Height          =   300
            Left            =   150
            TabIndex        =   23
            Text            =   "浏览……"
            Top             =   360
            Width           =   4170
         End
         Begin VB.ListBox lstFiles 
            Height          =   1680
            Left            =   120
            TabIndex        =   22
            Top             =   795
            Width           =   4185
         End
         Begin VB.Label Label8 
            Caption         =   "为防止失去响应，设置了一次最多批量转1000个txt页面"
            Height          =   375
            Left            =   600
            TabIndex        =   30
            Top             =   5280
            Width           =   5175
         End
         Begin VB.Label lblCurrentDir 
            Caption         =   "当前路径"
            Height          =   495
            Left            =   240
            TabIndex        =   26
            Top             =   2760
            Width           =   735
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Openwiki数据库转换成Dokuwiki的Txt存储库"
         Height          =   4935
         Left            =   -74640
         TabIndex        =   1
         Top             =   600
         Width           =   6015
         Begin VB.TextBox Text1 
            BackColor       =   &H8000000B&
            Height          =   375
            Left            =   5160
            TabIndex        =   13
            Top             =   1560
            Width           =   375
         End
         Begin VB.CommandButton Cmdasci 
            Caption         =   "查看输出文件"
            Height          =   495
            Left            =   1560
            TabIndex        =   12
            Top             =   4200
            Width           =   1215
         End
         Begin VB.CommandButton Command2 
            Caption         =   "选择表"
            Height          =   495
            Left            =   4080
            TabIndex        =   11
            ToolTipText     =   "Click to retrieve Tables stored in Database"
            Top             =   960
            Width           =   1455
         End
         Begin VB.ListBox List1 
            Height          =   1320
            Left            =   1560
            TabIndex        =   10
            TabStop         =   0   'False
            ToolTipText     =   "Select table to be exported  from list"
            Top             =   1560
            Width           =   2295
         End
         Begin VB.TextBox sqltxt 
            Height          =   735
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   9
            ToolTipText     =   "Skip the table Name as it will be picked from the table box above. (Ex:"" Select * from"" and NOT ""Select * from <tablename>)"
            Top             =   3360
            Width           =   5055
         End
         Begin VB.CommandButton Command1 
            Caption         =   "选择数据库"
            Height          =   495
            Left            =   4080
            TabIndex        =   8
            ToolTipText     =   "Pick Database "
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox tabtxt 
            Height          =   375
            Left            =   1800
            TabIndex        =   7
            Text            =   "openwiki_revisions"
            ToolTipText     =   "Table to be exported"
            Top             =   960
            Width           =   2055
         End
         Begin VB.TextBox loc 
            Height          =   375
            Left            =   1560
            TabIndex        =   6
            Text            =   "OpenWikiDist.mdb"
            Top             =   360
            Width           =   2295
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H8000000B&
            Enabled         =   0   'False
            Height          =   375
            Left            =   5160
            TabIndex        =   5
            ToolTipText     =   "Number of fields to be exported"
            Top             =   2040
            Width           =   375
         End
         Begin VB.CommandButton Cmdexit 
            Caption         =   "退出"
            Height          =   495
            Left            =   2760
            TabIndex        =   4
            Top             =   4200
            Width           =   1215
         End
         Begin VB.CommandButton Cmdexp 
            Caption         =   "开始转换"
            Height          =   495
            Left            =   360
            TabIndex        =   3
            Top             =   4200
            Width           =   1215
         End
         Begin VB.CommandButton Cmdgb2utf 
            Caption         =   "转换内码"
            Height          =   495
            Left            =   3960
            TabIndex        =   2
            Top             =   4200
            Width           =   1215
         End
         Begin MSComDlg.CommonDialog CmDialog1 
            Left            =   360
            Top             =   2040
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label6 
            Caption         =   "字段总数："
            Height          =   255
            Left            =   4080
            TabIndex        =   19
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label Label5 
            Caption         =   "表总清单："
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "SQL 命令行 (将忽略上面所选表)"
            ForeColor       =   &H80000001&
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   3000
            Width           =   3855
         End
         Begin VB.Label Label3 
            Caption         =   "待转化表："
            ForeColor       =   &H80000001&
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   1020
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "指定字段数："
            ForeColor       =   &H80000001&
            Height          =   255
            Left            =   4080
            TabIndex        =   15
            Top             =   2040
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "数据库路径："
            ForeColor       =   &H80000001&
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   480
            Width           =   1095
         End
      End
      Begin VB.Label Label7 
         Caption         =   $"Form1.frx":1D32
         ForeColor       =   &H80000001&
         Height          =   615
         Left            =   -74640
         TabIndex        =   21
         Top             =   5880
         Width           =   5895
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim fso As New FileSystemObject
Dim strDrive As String
Dim strFolder As String

Public sCodePage As Long    '用于UTF8编码和解码
Public cnvUni2 As String
Public cnvUni As String

Private Sub btnConvert_Click()
    'Call AppendMWikiXML(转换页面路径)
    Call AppendMWikiXML(txtCurrentDir.Text)
    'Call AppendMWikiXML("E:\User\桌面\wikiconverter\page")
    'FindFiles txtCurrentDir.Text, "", ""
End Sub

Private Sub btnDEUtf8_Click()
' 递归搜索文件直到无结果
Dim strFilePath, strFileTempPath, strFileTempName, strFileTempDir As String
Dim u As New cUTF8
Dim clsdb As New dbutils

Dim i As Integer
    Dim s As String
    Dim o_bytConent() As Byte



   i = 0
   frmMain.List2.Clear
   Do Until i = 1000   '本次待处理的为786个TXT文件
        If i < 0 Then Exit Do
        i = i + 1
        strFilePath = FindOnefile$(txtCurrentDir.Text, "*.txt")
        If strFilePath = "" Then Exit Do
        If Right$(strFilePath, 4) = ".txt" Then
        frmMain.List2.AddItem "文件 | " & ExtractFileName("" & strFilePath)
            's = strFilePath
            'u.DecodeURL s, o_bytConent
            's = StrConv(o_bytConent, vbUnicode)
            'strFileTempPath = u.DecodeUTF8(s)
                       
            strFileTempName = ExtractFileName("" & strFilePath)
            strFileTempDir = Replace(strFilePath, strFileTempName, "")
            strFileTempPath = strFileTempDir & clsdb.UTF8Decode(strFileTempName)
            strFileTempPath = Replace$(strFileTempPath, ".txt", ".bin")
            Name strFilePath As strFileTempPath
            strFilePath = strFileTempPath
        End If
        
        'If strFilePath <> "" And strFilePath <> "wikiconverter.exe" Then n = n + 1
      DoEvents
   Loop
   
' 把临时改为的bin扩展名文件恢复原来的TXT扩展名
   i = 0
   Do Until i = 1000   '本次待处理的为786个TXT文件
        If i < 0 Then Exit Do
        i = i + 1
        strFilePath = FindOnefile$(txtCurrentDir.Text, "*.bin")
        If strFilePath = "" Then Exit Do
        If Right$(strFilePath, 4) = ".bin" Then
            strFileTempPath = Replace$(strFilePath, ".bin", ".txt")
            Name strFilePath As strFileTempPath
            strFilePath = strFileTempPath
        End If
     DoEvents
   Loop
   

End Sub

Private Sub Cmdasci_Click()
Dim expfile As dbutils
Set expfile = New dbutils
Dim retval

retval = Shell("explorer.exe " & App.Path & "\" & "page\", 1)
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub Cmdexp_Click()
Dim err_handler
On Error GoTo err_handler
Dim tableName As String
Dim thepath
Dim mysql
Dim myfile
Dim mycounter As Variant
Dim counter1 As dbutils
Dim file1 As dbutils
Dim exp As dbutils
Dim skl  As dbutils
Dim mydata As dbutils

Set mydata = New dbutils
Set counter1 = New dbutils
Set file1 = New dbutils
Set exp = New dbutils
Set skl = New dbutils


'path of db
thepath = Trim(loc.Text)
mydata.thedbfilelocation = thepath

' Database table to query
tableName = Trim(tabtxt.Text)

'sql command
If sqltxt.Text <> "" Then
    mysql = sqltxt.Text & " " & tableName
Else
    mysql = "select * from " & tableName
End If
skl.thesql = mysql

'Name of ASCI file
myfile = "export.txt"
file1.thefile = myfile

'number of fields to be exported to the ASCI file
mycounter = Text2.Text
counter1.thecounter = mycounter


If loc.Text <> "" And tabtxt.Text <> "" And mycounter <> "" Then
   'And IsNumeric(Text2.Text) = True Then
    Call exp.exportmydb(thepath, mysql, myfile, mycounter)
Else
   MsgBox ("请点击'选择表'，并选定'openwiki_revisions'")

   Cmdexp.Enabled = False
   Cmdasci.Enabled = False
   Exit Sub
   
End If

Exit Sub
  
    

err_handler:
MsgBox (Err.Description)
Cmdexp.Enabled = False
Exit Sub

End Sub

Private Sub Cmdgb2utf_Click()
    Unload frmMain
    Load frmGb2utf
    frmGb2utf.Show
End Sub

Private Sub Command1_Click()
Dim openerror1
On Error GoTo openerror1
Cmdexp.Enabled = True
Cmdasci.Enabled = True
CmDialog1.Filter = "All Files (*.*)|*.*| Access Database Files (*.mdb)|*.mdb|  "
CmDialog1.FilterIndex = 2
CmDialog1.Action = 1
loc.Text = CmDialog1.FileName
Exit Sub
openerror1:
MsgBox (Err.Description)
Exit Sub
End Sub

Private Sub Command2_Click()

Dim err_handler
On Error GoTo err_handler

Dim thepath
Dim mydata As dbutils
Set mydata = New dbutils
Dim myfile As dbutils
Set myfile = New dbutils


thepath = Trim(loc.Text)
mydata.thedbfilelocation = thepath

Cmdexp.Enabled = True
Cmdasci.Enabled = True

If loc <> "" Then
    Call myfile.gettables(thepath, frmMain, List1)
    
Else
    MsgBox ("Select a DB first!")

End If



Exit Sub

err_handler:
MsgBox (Err.Description)
Exit Sub

End Sub



Private Sub Label8_Click()

End Sub

Private Sub List1_Click()

Dim thepath
Dim myfile As dbutils
Set myfile = New dbutils

Dim thetable
'selected item in list gets displayed in tabtxt
tabtxt.Text = List1.List(List1.ListIndex)

thepath = Trim(loc.Text)
thetable = Trim(tabtxt.Text)

'display max number of fields in table
Text1.Text = myfile.getfields(thepath, thetable) - 1
Text2.Text = myfile.getfields(thepath, thetable) - 1
End Sub

Private Sub loc_Change()
Cmdexp.Enabled = True
Cmdasci.Enabled = True
End Sub


Private Sub tabtxt_Change()
Cmdexp.Enabled = True
Cmdasci.Enabled = True
End Sub

Private Sub Text2_Change()
Cmdexp.Enabled = True
Cmdasci.Enabled = True
End Sub




Private Sub cmbDrives_Click()
    Dim drive As drive
    Dim File As File
    Dim SubFolder As Folder
    Dim i As Integer
    i = 0
    lstFiles.Clear
    If cmbDrives = "" Then Exit Sub
    strDrive = cmbDrives.Text
    strFolder = ""
    Set drive = fso.GetDrive(cmbDrives.Text)


    If drive.IsReady Then


        For Each File In drive.RootFolder.Files
            lstFiles.AddItem File.name, i
            i = i + 1
        Next
        i = lstFiles.ListCount


        For Each SubFolder In drive.RootFolder.SubFolders
            lstFiles.AddItem SubFolder, i
            i = i + 1
        Next
    Else
        MsgBox "Drives Not ready"
    End If
End Sub


Private Sub cmdup_Click()
    Dim Folder As Folder
    Dim File As File
    Dim SubFolder As Folder
    Dim i As Integer
    If strDrive = "" Then Exit Sub
    If strFolder = "" Then Exit Sub
    '当前文件夹
    Set Folder = fso.GetFolder(strDrive & strFolder)
    '上级文件夹
    strFolder = Left(strFolder, InStrRev(strFolder, "\") - 1)
    lstFiles.Clear
    txtCurrentDir.Text = cmbDrives.Text & strFolder
    

    If Not Folder.ParentFolder Is Nothing Then
        '添加文件


        For Each File In Folder.ParentFolder.Files
            lstFiles.AddItem File.name, i
            i = i + 1
        Next
        i = lstFiles.ListCount
        '添加子文件夹


        For Each SubFolder In Folder.ParentFolder.SubFolders
            lstFiles.AddItem SubFolder, i
            i = i + 1
        Next
    Else


        For Each File In Folder.Files
            lstFiles.AddItem File.name, i
            i = i + 1
        Next
        i = lstFiles.ListCount


        For Each SubFolder In Folder.SubFolders
            lstFiles.AddItem SubFolder, i
            i = i + 1
        Next
    End If
End Sub


Private Sub Form_Load()
    Dim drive As drive
    Dim i As Integer
    i = 0
    '添加驱动器


    For Each drive In fso.Drives
        cmbDrives.AddItem drive.Path, i
        i = i + 1
    Next
    
    
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Set fso = Nothing
End Sub


Private Sub lstFiles_Click()
    Dim Folder As Folder
    Dim SubFolder As Folder
    Dim File As File
    Dim i As Integer
    i = 0


    If Not lstFiles.SelCount > 1 Then

        If InStr(lstFiles.Text, ":\") Then
            Set Folder = fso.GetFolder(lstFiles.Text)
            lstFiles.Clear
            strFolder = strFolder & "\" & Folder.name
            txtCurrentDir.Text = cmbDrives.Text & strFolder
            
            '添加全部文件


            For Each File In Folder.Files
                lstFiles.AddItem File.name, i
                i = i + 1
            Next
            i = lstFiles.ListCount
            '添加子文件夹


            For Each SubFolder In Folder.SubFolders
                lstFiles.AddItem SubFolder, i
                i = i + 1
            Next
        End If
    End If
End Sub


