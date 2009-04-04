VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Openwiki2Dokuwiki"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   6720
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Caption         =   "Openwiki数据库转换成Dokuwiki的Txt存储库"
      Height          =   4935
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   6015
      Begin VB.CommandButton Cmdgb2utf 
         Caption         =   "转换内码"
         Height          =   495
         Left            =   3960
         TabIndex        =   18
         Top             =   4200
         Width           =   1215
      End
      Begin VB.CommandButton Cmdexp 
         Caption         =   "开始转换"
         Height          =   495
         Left            =   360
         TabIndex        =   11
         Top             =   4200
         Width           =   1215
      End
      Begin VB.CommandButton Cmdexit 
         Caption         =   "退出"
         Height          =   495
         Left            =   2760
         TabIndex        =   10
         Top             =   4200
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         Height          =   375
         Left            =   5160
         TabIndex        =   9
         ToolTipText     =   "Number of fields to be exported"
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox loc 
         Height          =   375
         Left            =   1560
         TabIndex        =   8
         Text            =   "OpenWikiDist.mdb"
         Top             =   360
         Width           =   2295
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
      Begin VB.CommandButton Command1 
         Caption         =   "选择数据库"
         Height          =   495
         Left            =   4080
         TabIndex        =   6
         ToolTipText     =   "Pick Database "
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox sqltxt 
         Height          =   735
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   5
         ToolTipText     =   "Skip the table Name as it will be picked from the table box above. (Ex:"" Select * from"" and NOT ""Select * from <tablename>)"
         Top             =   3360
         Width           =   5055
      End
      Begin VB.ListBox List1 
         Height          =   1320
         Left            =   1560
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Select table to be exported  from list"
         Top             =   1560
         Width           =   2295
      End
      Begin VB.CommandButton Command2 
         Caption         =   "选择表"
         Height          =   495
         Left            =   4080
         TabIndex        =   3
         ToolTipText     =   "Click to retrieve Tables stored in Database"
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton Cmdasci 
         Caption         =   "查看输出文件"
         Height          =   495
         Left            =   1560
         TabIndex        =   2
         Top             =   4200
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000B&
         Height          =   375
         Left            =   5160
         TabIndex        =   1
         Top             =   1560
         Width           =   375
      End
      Begin MSComDlg.CommonDialog CmDialog1 
         Left            =   360
         Top             =   2040
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label1 
         Caption         =   "数据库路径："
         ForeColor       =   &H80000001&
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "指定字段数："
         ForeColor       =   &H80000001&
         Height          =   255
         Left            =   4080
         TabIndex        =   16
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "待转化表："
         ForeColor       =   &H80000001&
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1020
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "SQL 命令行 (将忽略上面所选表)"
         ForeColor       =   &H80000001&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   3000
         Width           =   3855
      End
      Begin VB.Label Label5 
         Caption         =   "表总清单："
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "字段总数："
         Height          =   255
         Left            =   4080
         TabIndex        =   12
         Top             =   1560
         Width           =   975
      End
   End
   Begin VB.Label Label7 
      Caption         =   $"Form1.frx":1CFA
      ForeColor       =   &H80000001&
      Height          =   615
      Left            =   360
      TabIndex        =   19
      Top             =   5280
      Width           =   5895
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


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

