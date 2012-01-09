Attribute VB_Name = "MFindFile"
Option Explicit

Const ATTR_DIRECTORY = 16

Public Function FindOnefile$(srcPath$, Filename$)

' Description
'     Finds a file in the subdirs under srcPath
'
' Parameters
'     Name                 Type     Value
'     ------------------------------------------------------------
'     srcPath              String   The path so start searching in
'     Filename             String   The file to look for
'
' Returns
'     The first occurence of Filename$
'
' Last modified by Jens Balchen 10.03.1996
' $为字符串型类型标识符，主要在未声明时使用
' DIR函数支持通配符*,?，但它不能无条件递归调用

Dim DirReturn As String, I As Integer
Dim subdirs$
Dim filefound$
Dim path$

On Error GoTo errHandler

   ' If Path lacks a "\", add one to the end
   If Right$(srcPath, 1) <> "\" Then srcPath = srcPath & "\"
   srcPath = UCase$(srcPath)
   Filename$ = UCase$(Filename$)
   
   '如果文件名为空，则把路径名也清为空
   If Filename$ = "" Then srcPath = ""
   
   ' Now find if the file is here
   DirReturn = Dir(srcPath & Filename, 0)

   ' If it is, return the filename and dir
   If DirReturn$ <> "" Then
      'FindOnefile$ = srcPath$ & Filename$
      FindOnefile$ = srcPath$ & DirReturn$
      Exit Function
   End If

   ' No, it wasn't here. Do all the subdirs
   ' Initialize var to hold filenames
   DirReturn = Dir(srcPath & "*.*", ATTR_DIRECTORY)
   
   ' Find all subdirs
   Do While DirReturn <> ""
      ' Make sure we don't do anything with "." and "..", they aren't real files
      If DirReturn <> "." And DirReturn <> ".." Then
         DoEvents
         If GetAttr(srcPath & DirReturn) = ATTR_DIRECTORY Then
            ' It's a dir. Add it to dirlist
            subdirs$ = subdirs$ & srcPath & DirReturn & ";"
         End If
      End If
      DirReturn = Dir
   Loop

   ' Do all subs
   Do
      
      I% = I% + 1
      If Mid$(subdirs$, I%, 1) = ";" Then
         path$ = Left$(subdirs$, I% - 1)
         If path$ <> "" Then
            filefound$ = FindOnefile$(path$, Filename$)
            If filefound$ <> "" Then
               FindOnefile$ = filefound$
               Exit Function
            End If
            subdirs$ = Right$(subdirs$, Len(subdirs$) - I%)
            I% = 0
         End If
      End If

      DoEvents

      If subdirs$ = "" Then Exit Do

   Loop
   
GoTo ExitFunc
   
errHandler:

    MsgBox "找不到指定文件: " & Err.Number & vbCrLf & Err.Description
       
ExitFunc:

   Exit Function

End Function

