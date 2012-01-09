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
'UTF8 编码/解码
'--------------------API声明部分--------------------


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
Dim I As Integer
Dim recproc As Integer
Set thedb = OpenDatabase(thedbfilelocation)
Set myrecord = thedb.OpenRecordset(thesql)
 

'check that user entered right number of fields
If CInt(thecounter) > myrecord.Fields.count Then
    MsgBox ("数据库中没那么多的字段")
 Close filenum
 Exit Function
Else

    While Not myrecord.EOF
        recproc = recproc + 1
        
        'If myrecord.Fields("wrv_name") = "Google搜索帮助" Then
        
        filenum = FreeFile
        thefile = UTF8Encode(LCase(myrecord.Fields("wrv_name"))) & ".txt"
        Open App.path & "\page\" & thefile For Output As filenum ' to clear file
        Close filenum
        Open App.path & "\page\" & thefile For Append As filenum
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
Dim I

'clear list
mylist.Visible = True
mylist.Clear


'specify DB
Set thedb = OpenDatabase(thedbfilelocation)


With thedb

For I = 0 To .TableDefs.count - 1
    If Left$(.TableDefs(I).name, 4) <> "MSys" Then
    mylist.AddItem .TableDefs(I).name
    End If
Next I

End With

Exit Function

err_2:
MsgBox Err.Description
Exit Function


End Function

Public Function getfields(thedbfilelocation, thesql)


Dim err_2
On Error GoTo err_3
Dim I


'specify DB
Set thedb = OpenDatabase(thedbfilelocation)
Set myrec = thedb.OpenRecordset(thesql)

getfields = myrec.Fields.count

Exit Function

err_3:
MsgBox Err.Description
Exit Function


End Function


'进行UrlEncode编码
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




'UTF-8编码
 Public Function UTF8Encode(ByVal szInput As String) As String
    Dim wch  As String
    Dim uch As String
    Dim szRet As String
    Dim x As Long
    Dim inputLen As Long
    Dim nAsc  As Long
    Dim nAsc2 As Long
    Dim nAsc3 As Long
    
    If szInput = "" Then
        UTF8Encode = szInput
        Exit Function
    End If
    inputLen = Len(szInput)
    For x = 1 To inputLen
    '得到每个字符
        wch = Mid(szInput, x, 1)
        '得到相应的UNICODE编码
        nAsc = AscW(wch)
    '对于<0的编码　其需要加上65536
        If nAsc < 0 Then nAsc = nAsc + 65536
    '对于<128位的ASCII的编码则无需更改
        If (nAsc And &HFF80) = 0 Then
            szRet = szRet & wch
        Else
            If (nAsc And &HF000) = 0 Then
            '真正的第二层编码范围为000080 - 0007FF
            'Unicode在范围D800-DFFF中不存在任何字符，基本多文种平面中约定了这个范围用于UTF-16扩展标识辅助平面（两个UTF-16表示一个辅助平面字符）.
            '当然，任何编码都是可以被转换到这个范围，但在unicode中他们并不代表任何合法的值。
    
                uch = "%" & Hex(((nAsc \ 2 ^ 6)) Or &HC0) & Hex(nAsc And &H3F Or &H80)
                szRet = szRet & uch
                
            Else
            '第三层编码00000800 – 0000FFFF
            '首先取其前四位与11100000进行或去处得到UTF-8编码的前8位
            '其次取其前10位与111111进行并运算，这样就能得到其前10中最后6位的真正的编码　再与10000000进行或运算来得到UTF-8编码中间的8位
            '最后将其与111111进行并运算，这样就能得到其最后6位的真正的编码　再与10000000进行或运算来得到UTF-8编码最后8位编码
                uch = "%" & Hex((nAsc \ 2 ^ 12) Or &HE0) & "%" & _
                Hex((nAsc \ 2 ^ 6) And &H3F Or &H80) & "%" & _
                Hex(nAsc And &H3F Or &H80)
                szRet = szRet & uch
            End If
        End If
    Next
    
    UTF8Encode = szRet
End Function

'UTF-8解码(3-12更改,可以解多个字符串 可供正常使用)
'http://topic.csdn.net/u/20100215/09/575b4408-4cd0-422b-987c-74b7112b9725.html
  
 Public Function UTF8Decode(ByVal code As String) As String
     If code = "" Then
         UTF8Decode = ""
         Exit Function
     End If
     
     Dim tmp As String
     Dim decodeStr As String
     Dim codelen As Long
     Dim result As String
     Dim leftStr As String
      
     leftStr = Left(code, 1)
      
     While (code <> "")
         codelen = Len(code)
         leftStr = Left(code, 1)
         If leftStr = "%" Then
                 If (Mid(code, 2, 1) = "C" Or Mid(code, 2, 1) = "B") Then
                     decodeStr = Replace(Mid(code, 1, 6), "%", "")
                     tmp = c10ton(Val("&H" & Hex(Val("&H" & decodeStr) And &H1F3F)))
                     tmp = String(16 - Len(tmp), "0") & tmp
                     UTF8Decode = UTF8Decode & UTF8Decode & ChrW(Val("&H" & c2to16(Mid(tmp, 3, 4)) & c2to16(Mid(tmp, 7, 2) & Mid(tmp, 11, 2)) & Right(decodeStr, 1)))
                     code = Right(code, codelen - 6)
                 ElseIf (Mid(code, 2, 1) = "E") Then
                     decodeStr = Replace(Mid(code, 1, 9), "%", "")
                     tmp = c10ton((Val("&H" & Mid(Hex(Val("&H" & decodeStr) And &HF3F3F), 2, 3))))
                     tmp = String(10 - Len(tmp), "0") & tmp
                     UTF8Decode = UTF8Decode & ChrW(Val("&H" & (Mid(decodeStr, 2, 1) & c2to16(Mid(tmp, 1, 4)) & c2to16(Mid(tmp, 5, 2) & Right(tmp, 2)) & Right(decodeStr, 1))))
                     code = Right(code, codelen - 9)
                 End If
         Else
             UTF8Decode = UTF8Decode & leftStr
             code = Right(code, codelen - 1)
         End If
     Wend
 End Function
 
 '10进制转n进制(默认2)
Public Function c10ton(ByVal x As Integer, Optional ByVal n As Integer = 2) As String
    Dim I As Integer
    I = x \ n
    If I > 0 Then
        If x Mod n > 10 Then
            c10ton = c10ton(I, n) + Chr(x Mod n + 55)
        Else
            c10ton = c10ton(I, n) + CStr(x Mod n)
        End If
    Else
        If x > 10 Then
            c10ton = Chr(x + 55)
        Else
            c10ton = CStr(x)
        End If
    End If
End Function

'二进制代码转换为十六进制代码
Public Function c2to16(ByVal x As String) As String
   Dim I As Long
   I = 1
   For I = 1 To Len(x) Step 4
      c2to16 = c2to16 & Hex(c2to10(Mid(x, I, 4)))
   Next
End Function


'二进制代码转换为十进制代码
Public Function c2to10(ByVal x As String) As String
   c2to10 = 0
   If x = "0" Then Exit Function
   Dim I As Long
   I = 0
   For I = 0 To Len(x) - 1
      If Mid(x, Len(x) - I, 1) = "1" Then c2to10 = c2to10 + 2 ^ (I)
   Next
End Function


