Attribute VB_Name = "modTxtUtf"
Public txtfilesize As Long


'读取UTF-8编码文件测试
' Text1.Text = File_get_contents("c:\1.txt", "UTF-8")
'参数: Path 文件路径
'      Unicode 文件编码
Public Function File_get_contents(Path As String, Optional Unicode = "GB2312")
    Dim arrBinary() As Byte
    txtfilesize = 0
    Open Path For Binary As #1
        txtfilesize = LOF(1)
        ReDim arrBinary(txtfilesize - 1)
        Get #1, , arrBinary()
    Close #1
    File_get_contents = BytesToBstr(arrBinary, Unicode)
End Function
Public Function BytesToBstr(Binary, Unicode)
     Dim objstream As Object
     Set objstream = CreateObject("ADODB.Stream")
     objstream.Type = 1
     objstream.Mode = 3
     objstream.Open
     objstream.Write Binary
     objstream.Position = 0
     objstream.Type = 2
     objstream.Charset = Unicode
     BytesToBstr = objstream.ReadText
     objstream.Close
End Function
