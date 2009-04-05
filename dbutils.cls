VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dbutils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mysql As String
Private dbfilepath As String
Private exportfile
Private mycounter As Variant

'---------------------------------------------------
'UTF8 ����/����
'--------------------API��������--------------------


Public Property Get thedbfilelocation() As String
thedbfilelocation = dbfilepath
End Property

Public Property Let thedbfilelocation(ByVal vNewValue As String)
dbfilepath = vNewValue
End Property

Public Property Get thesql() As Variant
thesql = mysql
End Property

Public Property Let thesql(ByVal vNewValue As Variant)
mysql = vNewValue
End Property

Public Property Get thecounter() As Variant
thecounter = mycounter
End Property

Public Property Let thecounter(ByVal vNewValue As Variant)

mycounter = vNewValue
End Property

Public Property Get thefile() As Variant
thefile = exportfile
End Property

Public Property Let thefile(ByVal vNewValue As Variant)
exportfile = vNewValue
End Property



Public Function exportmydb(thedbfilelocation, thesql, thefile, thecounter)
'#############################################################
'# this class function exports a database to an ASCI file    #
'# Author: Fettah Hounain                                    #
'# email: is09@netscape.net (0 is a zero not an o)           #
'#############################################################
Dim err_1
On Error GoTo err_1
Dim filenum
Dim i As Integer
Dim recproc As Integer
Set thedb = OpenDatabase(thedbfilelocation)
Set myrecord = thedb.OpenRecordset(thesql)
 

'check that user entered right number of fields
If CInt(thecounter) > myrecord.Fields.count Then
    MsgBox ("���ݿ���û��ô����ֶ�")
 Close filenum
 Exit Function
Else

    While Not myrecord.EOF
        recproc = recproc + 1
        
        'If myrecord.Fields("wrv_name") = "Google��������" Then
        
        filenum = FreeFile
        thefile = UTF8Encode(LCase(myrecord.Fields("wrv_name"))) & ".txt"
        Open App.Path & "\page\" & thefile For Output As filenum ' to clear file
        Close filenum
        Open App.Path & "\page\" & thefile For Append As filenum
        Print #filenum, (myrecord.Fields("wrv_text"));   ' Output text.
        Close filenum   ' Close file.
        Debug.Print recproc
        Debug.Print myrecord.Fields("wrv_name")
       'End If
    myrecord.MoveNext
    Wend

End If

myrecord.Close
thedb.Close


Exit Function
err_1:
MsgBox (Err.Description)
Close filenum
Exit Function

End Function

Public Function gettables(thedbfilelocation, myform As Form, mylist As ListBox)
Dim err_2
On Error GoTo err_2
Dim i

'clear list
mylist.Visible = True
mylist.Clear


'specify DB
Set thedb = OpenDatabase(thedbfilelocation)


With thedb

For i = 0 To .TableDefs.count - 1
    If Left$(.TableDefs(i).name, 4) <> "MSys" Then
    mylist.AddItem .TableDefs(i).name
    End If
Next i

End With

Exit Function

err_2:
MsgBox Err.Description
Exit Function


End Function

Public Function getfields(thedbfilelocation, thesql)


Dim err_2
On Error GoTo err_3
Dim i


'specify DB
Set thedb = OpenDatabase(thedbfilelocation)
Set myrec = thedb.OpenRecordset(thesql)

getfields = myrec.Fields.count

Exit Function

err_3:
MsgBox Err.Description
Exit Function


End Function


'����UrlEncode����
Public Function URLEncode(StringToEncode As String, Optional _
   UsePlusRatherThanHexForSpace As Boolean = False) As String

Dim TempAns As String
Dim CurChr As Integer
CurChr = 1
Do Until CurChr - 1 = Len(StringToEncode)
  Select Case Asc(Mid(StringToEncode, CurChr, 1))
    Case 48 To 57, 65 To 90, 97 To 122
      TempAns = TempAns & Mid(StringToEncode, CurChr, 1)
    Case 32
      If UsePlusRatherThanHexForSpace = True Then
        TempAns = TempAns & "+"
      Else
        TempAns = TempAns & "%" & Hex(32)
      End If
   Case Else
         TempAns = TempAns & "%" & _
              Format(Hex(Asc(Mid(StringToEncode, _
              CurChr, 1))), "00")
End Select

  CurChr = CurChr + 1
Loop

URLEncode = TempAns
End Function

' URLDecode function in Perl for reference
' both VB and Perl versions must return same
'
' sub urldecode{
'  local($val)=@_;
'  $val=~s/\+/ /g;
'  $val=~s/%([0-9A-H]{2})/pack('C',hex($1))/ge;
'  return $val;
' }

Public Function URLDecode(StringToDecode As String) As String

Dim TempAns As String
Dim CurChr As Integer

CurChr = 1

Do Until CurChr - 1 = Len(StringToDecode)
  Select Case Mid(StringToDecode, CurChr, 1)
    Case "+"
      TempAns = TempAns & " "
    Case "%"
      TempAns = TempAns & Chr(Val("&h" & _
         Mid(StringToDecode, CurChr + 1, 2)))
       CurChr = CurChr + 2
    Case Else
      TempAns = TempAns & Mid(StringToDecode, CurChr, 1)
  End Select

CurChr = CurChr + 1
Loop

URLDecode = TempAns
End Function




'UTF-8����
 Public Function UTF8Encode(ByVal szInput As String) As String
    Dim wch  As String
    Dim uch As String
    Dim szRet As String
    Dim X As Long
    Dim inputLen As Long
    Dim nAsc  As Long
    Dim nAsc2 As Long
    Dim nAsc3 As Long
    
    If szInput = "" Then
        UTF8Encode = szInput
        Exit Function
    End If
    inputLen = Len(szInput)
    For X = 1 To inputLen
    '�õ�ÿ���ַ�
        wch = Mid(szInput, X, 1)
        '�õ���Ӧ��UNICODE����
        nAsc = AscW(wch)
    '����<0�ı��롡����Ҫ����65536
        If nAsc < 0 Then nAsc = nAsc + 65536
    '����<128λ��ASCII�ı������������
        If (nAsc And &HFF80) = 0 Then
            szRet = szRet & wch
        Else
            If (nAsc And &HF000) = 0 Then
            '�����ĵڶ�����뷶ΧΪ000080 - 0007FF
            'Unicode�ڷ�ΧD800-DFFF�в������κ��ַ�������������ƽ����Լ���������Χ����UTF-16��չ��ʶ����ƽ�棨����UTF-16��ʾһ������ƽ���ַ���.
            '��Ȼ���κα��붼�ǿ��Ա�ת���������Χ������unicode�����ǲ��������κκϷ���ֵ��
    
                uch = "%" & Hex(((nAsc \ 2 ^ 6)) Or &HC0) & Hex(nAsc And &H3F Or &H80)
                szRet = szRet & uch
                
            Else
            '���������00000800 �C 0000FFFF
            '����ȡ��ǰ��λ��11100000���л�ȥ���õ�UTF-8�����ǰ8λ
            '���ȡ��ǰ10λ��111111���в����㣬�������ܵõ���ǰ10�����6λ�������ı��롡����10000000���л��������õ�UTF-8�����м��8λ
            '�������111111���в����㣬�������ܵõ������6λ�������ı��롡����10000000���л��������õ�UTF-8�������8λ����
                uch = "%" & Hex((nAsc \ 2 ^ 12) Or &HE0) & "%" & _
                Hex((nAsc \ 2 ^ 6) And &H3F Or &H80) & "%" & _
                Hex(nAsc And &H3F Or &H80)
                szRet = szRet & uch
            End If
        End If
    Next
    
    UTF8Encode = szRet
End Function
