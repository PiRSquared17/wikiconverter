VERSION 5.00
Begin VB.Form frmText 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "定义替换规则"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5505
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdExit 
      Caption         =   "返回(&Y)"
      Height          =   300
      Left            =   4451
      TabIndex        =   4
      Top             =   210
      Width           =   800
   End
   Begin VB.ComboBox cmbAction 
      Height          =   300
      Left            =   3284
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   210
      Width           =   1000
   End
   Begin VB.ComboBox cmbNumber 
      Height          =   300
      Left            =   1259
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   210
      Width           =   1000
   End
   Begin VB.TextBox txtNew 
      Height          =   900
      Left            =   252
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   2715
      Width           =   5000
   End
   Begin VB.TextBox txtOld 
      Height          =   900
      Left            =   252
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1255
      Width           =   5000
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   -15
      X2              =   5585
      Y1              =   735
      Y2              =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   5610
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "操作(&O)："
      Height          =   180
      Left            =   2429
      TabIndex        =   2
      Top             =   270
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "规则号(&N)："
      Height          =   180
      Left            =   254
      TabIndex        =   0
      Top             =   270
      Width           =   990
   End
   Begin VB.Label lblNew 
      AutoSize        =   -1  'True
      Caption         =   "新文本："
      Height          =   180
      Left            =   252
      TabIndex        =   8
      Top             =   2445
      Width           =   720
   End
   Begin VB.Label lblOld 
      AutoSize        =   -1  'True
      Caption         =   "原文本："
      Height          =   180
      Left            =   252
      TabIndex        =   7
      Top             =   990
      Width           =   720
   End
End
Attribute VB_Name = "frmText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbAction_Click()
    If cmbAction.ListIndex > 0 Then
        txtOld.Enabled = False
        txtOld.Text = "<插入或追加新内容，这里不需要输入>"
    Else
        txtOld.Enabled = True
    End If
End Sub

Private Sub cmbNumber_Click()
    Dim idx1 As Integer, idx2 As Integer
    If cmbNumber.Tag <> cmbNumber.Text Then
        '改变条目
        If cmbNumber.Tag = "" Then
            idx1 = 0
        Else
            idx1 = CInt(cmbNumber.Tag)
        End If
        idx2 = CInt(cmbNumber.Text)
        
        If idx1 > 0 Then saveOption idx1
        
        showOption idx2
        
        cmbNumber.Tag = cmbNumber.Text
    End If
End Sub

Private Sub saveOption(ByVal idx As Integer)
    '保存规则到数组
    FileTextAction(idx - 1).TextAct = cmbAction.ListIndex
    FileTextAction(idx - 1).OldText = txtOld.Text
    FileTextAction(idx - 1).NewText = txtNew.Text
End Sub

Private Sub showOption(ByVal idx As Integer)
    '从数组读取规则
    cmbAction.ListIndex = FileTextAction(idx - 1).TextAct
    txtOld.Text = FileTextAction(idx - 1).OldText
    txtNew.Text = FileTextAction(idx - 1).NewText
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    cmbAction.AddItem "替换"
    cmbAction.AddItem "插入头"
    cmbAction.AddItem "追加尾"
    'cmbAction.ListIndex = 0

    For i = 1 To 9
        cmbNumber.AddItem CStr(i)
    Next
    cmbNumber.ListIndex = TextActionIdx
End Sub

Private Sub Form_Unload(Cancel As Integer)
    TextActionIdx = CInt(cmbNumber.Text) - 1
    saveOption CInt(cmbNumber.Text)
End Sub
