Attribute VB_Name = "modXMLdeal"
Public Sub AppendWikiXML(strRootFolder As String, strFile As String)

Dim m_AppPath As String
Dim xml_document As DOMDocument
Dim values_node As IXMLDOMNode
Dim values_node2 As IXMLDOMNode
Dim Root_Node As IXMLDOMNode
Dim n As Integer

    '��ȡ����·��
    m_AppPath = App.Path
    If Right$(m_AppPath, 1) <> "\" Then m_AppPath = m_AppPath & "\"
    '�����ĵ�
    Set xml_document = New DOMDocument
    xml_document.Load m_AppPath & "export2.xml"
    '����������κ�Ԫ�أ����˳�
    If xml_document.documentElement Is Nothing Then
        Exit Sub
    End If
    'Root_Node��Ϊ���ڵ�
    Set Root_Node = xml_document.selectSingleNode("mediawiki")
    '��ȡTotal�ڵ��Text
    Set values_node = Root_Node.selectSingleNode("siteinfo")
    Set values_node2 = values_node.selectSingleNode("sitename")
    'For N = 1 To CInt(values_node.Text)
        'Set values_node = Root_Node.selectSingleNode("_" & N)
        'With Text1
            '.Text = .Text + "title:" & GetNodeValue(values_node, "title", "???") & vbCrLf
            '.Text = .Text + "FileName:" & GetNodeValue(values_node, "FileName", "???") & vbCrLf
            '.Text = .Text + "Del:" & GetNodeValue(values_node, "Del", "???") & vbCrLf
            '.Text = .Text + "-------------------------------" & vbCrLf
        'End With
    'Next
    
    '����µĽڵ�
    
    'N = CInt(GetNodeValue(Root_Node, "Total", "?")) + 1
    '�����Ԫ��
    CreateNode 4, Root_Node, "page", ""
    Set values_node = Root_Node.selectSingleNode("Page")
    CreateNode 4, values_node, "Name", "����"
    CreateNode 4, values_node, "FileName", "2006"
    CreateNode 4, values_node, "Del", "0"
    '�޸�Totalֵ
    'Set values_node = Root_Node.selectSingleNode("Total")
    'values_node.Text = N
    xml_document.save m_AppPath & "export2.xml"

End Sub

Private Function GetNodeValue(ByVal start_at_node As IXMLDOMNode, _
                               ByVal node_name As String, _
                               Optional ByVal default_value As String = "") _
                               As String
Dim value_node As IXMLDOMNode
Set value_node = start_at_node.selectSingleNode(".//" & node_name)
If value_node Is Nothing Then
    GetNodeValue = default_value
Else
    GetNodeValue = value_node.Text
End If
End Function

Private Sub CreateNode(ByVal indent As Integer, _
ByVal parent As IXMLDOMNode, ByVal node_name As String, _
ByVal node_value As String)
Dim new_node As IXMLDOMNode


' Indent.
parent.appendChild parent.ownerDocument.createTextNode(Space$(indent))
' Create the new node.
Set new_node = parent.ownerDocument.createElement(node_name)
' Set the node's text value.
new_node.Text = node_value
' Add the node to the parent.
parent.appendChild new_node
' Add a new line.
parent.appendChild parent.ownerDocument.createTextNode(vbCrLf)
End Sub


'*************************************************************************
'Function:          AppendMWikiXML
'Desctiption:      ����һ��MEDIAWIKI��ʽ��Xml
'                       ��ʽ����:����
'                       <?xml version="1.0" encoding="UTF-8"?>
'                        <mediawiki>
'                            <page>
'                            <title>Hello world</title>
'                            <id>2</id>
'                            <revision>
'                              <id>12</id>
'                              <timestamp>2011-12-18T01:03:26Z</timestamp>
'                              <contributor>
'                                <username>Dragon</username>
'                                <id>1</id>
'                              </contributor>
'                              <comment>��"hello world"Ϊ���ݴ���ҳ��</comment>
'                              <text xml:space="preserve" bytes="11">hello world</text>
'                            </revision>
'                          </page>
'                        </mediawiki>
'
'Parameters:       strPagePath
'
'Return Value:      ���ض�Ӧ��xml�ęn
'*************************************************************************
'DATE               NAME                    DESCRIPTION
'-------------------------------------------------------------------------
'2012-01-07      QingHui Chen               Function create
'*************************************************************************
Public Function AppendMWikiXML(ByVal strPagePath As String) As String
                                               
On Error GoTo ErrHandle:

    Dim tempdoc As MSXML2.DOMDocument               '�����xml�ļ�����
    Dim EL_curElement As MSXML2.IXMLDOMElement      '������ڵ�
    Dim RootNode As MSXML2.IXMLDOMNode
    Dim l_IXMPI As IXMLDOMProcessingInstruction
    Dim l_strxml As String                          '���ɵ�XML�ęnString
   
   
    Dim node_Parent As MSXML2.IXMLDOMNode '���ڵ����
    Dim node_son As MSXML2.IXMLDOMNode '�ӽڵ����
    Dim node_grandson As MSXML2.IXMLDOMNode '��ڵ����
    Dim node_greatgrandson As MSXML2.IXMLDOMNode '����ڵ����
    Dim tempAttribute As IXMLDOMElement
    
   
    Dim i As Integer                                        '��ʱ����
   
    If Len(strPagePath) = 0 Then
        AppendMWikiXML = ""
        Exit Function
    End If
   
    Set tempdoc = New MSXML2.DOMDocument
    Set EL_curElement = tempdoc.createElement("mediawiki")     '����mediawiki��Ϊ���ӵ�
   
    Set tempdoc.documentElement = EL_curElement
    
    
    
' �ݹ������ļ�ֱ���޽��
   i = 0
   frmMain.List2.Clear
      Do Until i = 1000   '���δ������Ϊ786��TXT�ļ�
        If i < 0 Then Exit Do
        'I = I + 1
        strFilePath = FindOnefile$(strPagePath, "*.txt")
        If strFilePath = "" Then Exit Do
        
        If Right$(strFilePath, 4) = ".txt" Then
            strFileTempPath = Replace$(strFilePath, ".txt", ".bin")
            Name strFilePath As strFileTempPath
            strFilePath = strFileTempPath
            frmMain.List2.AddItem "�ļ� | " & ExtractFileName("" & strFilePath)
                 Set node_Parent = tempdoc.CreateNode(MSXML2.NODE_ELEMENT, "page", "")  '����page�ڵ�
                
                 EL_curElement.appendChild node_Parent   '��mediawiki�¼��� page �ڵ�
                 
                 Set node_son = tempdoc.CreateNode(MSXML2.NODE_ELEMENT, "title", "")  '���� title �ڵ�
                 node_Parent.appendChild node_son   '�� page �¼��� title �ڵ�
                 node_son.Text = ExtractFileName("" & strFilePath)   '�趨 title �ڵ�ֵ
                 node_son.Text = Replace(node_son.Text, ".bin", "")

                 
                 Set node_son = tempdoc.CreateNode(MSXML2.NODE_ELEMENT, "id", "")  '���� id �ڵ�
                 node_Parent.appendChild node_son   '�� page �¼��� id �ڵ�
                 'If strFilePath <> "" And strFilePath <> "wikiconverter.exe" Then i = i + 1
                 i = i + 1
                 node_son.Text = i   '�趨 id �ڵ�ֵ
                 
                 Set node_son = tempdoc.CreateNode(MSXML2.NODE_ELEMENT, "revision", "")  '���� revision �ڵ�
                 node_Parent.appendChild node_son   '�� page �¼��� revision �ڵ�
                 
                 Set node_grandson = tempdoc.CreateNode(MSXML2.NODE_ELEMENT, "id", "")  '���� revision id �ڵ�
                 node_son.appendChild node_grandson   '�� revision �¼��� id �ڵ�
                 node_grandson.Text = "1"     '�趨 revision id �ڵ�ֵ
                 Set node_grandson = tempdoc.CreateNode(MSXML2.NODE_ELEMENT, "timestamp", "")  '���� revision timestamp �ڵ�
                 node_son.appendChild node_grandson   '�� revision �¼��� timestamp �ڵ�
                 node_grandson.Text = "2012-01-08T08:08:28Z"  '�趨 revision timestamp �ڵ�ֵ
                 Set node_grandson = tempdoc.CreateNode(MSXML2.NODE_ELEMENT, "contributor", "")  '���� revision contributor �ڵ�
                 node_son.appendChild node_grandson   '�� revision �¼��� contributor �ڵ�
                 
                 Set node_greatgrandson = tempdoc.CreateNode(MSXML2.NODE_ELEMENT, "username", "")  '���� revision contributor username �ڵ�
                 node_grandson.appendChild node_greatgrandson   '�� contributor �¼���  revision contributor username  �ڵ�
                 node_greatgrandson.Text = "Dragon"  '�趨 revision contributor username �ڵ�ֵ
                 Set node_greatgrandson = tempdoc.CreateNode(MSXML2.NODE_ELEMENT, "id", "")  '���� revision contributor id �ڵ�
                 node_grandson.appendChild node_greatgrandson   '�� contributor �¼���  revision contributor id  �ڵ�
                 node_greatgrandson.Text = "2"  '�趨 revision contributor id �ڵ�ֵ
                 
                 Set node_grandson = tempdoc.CreateNode(MSXML2.NODE_ELEMENT, "comment", "")  '���� revision comment �ڵ�
                 node_son.appendChild node_grandson   '�� revision �¼��� comment �ڵ�
                 node_grandson.Text = "from Dokuwiki"  '�趨 revision contributor comment �ڵ�ֵ
                 
                 Set node_grandson = tempdoc.CreateNode(MSXML2.NODE_ELEMENT, "text", "")  '���� revision text �ڵ�
                 node_son.appendChild node_grandson   '�� revision �¼��� text �ڵ�
                 
                 node_grandson.Text = File_get_contents("" & strFilePath, "UTF-8")  '�趨 revision text �ڵ�ֵ

                 '<<---------���� revision text��� Attribute ���Լ��趨ֵ--------------
                 
                 Set tempAttribute = node_grandson
                 tempAttribute.setAttribute "xml:space", "preserve"
                 tempAttribute.setAttribute "bytes", txtfilesize

        End If
        
      DoEvents
   Loop
   
   
    
' ����ʱ��Ϊ��bin��չ���ļ��ָ�ԭ����TXT��չ��
   i = 0
   Do Until i = 1000   '���δ������Ϊ786��TXT�ļ�
        If i < 0 Then Exit Do
        i = i + 1
        strFilePath = FindOnefile$(strPagePath, "*.bin")
        If strFilePath = "" Then Exit Do
        If Right$(strFilePath, 4) = ".bin" Then
            strFileTempPath = Replace$(strFilePath, ".bin", ".txt")
            Name strFilePath As strFileTempPath
            strFilePath = strFileTempPath
        End If
     DoEvents
   Loop

   
   
    Set l_IXMPI = tempdoc.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")
    Call tempdoc.insertBefore(l_IXMPI, tempdoc.childNodes(0))
    l_strxml = tempdoc.xml
   
    '������һ�����޳ɹ�����
   tempdoc.save (strPagePath & "Export.xml")
   AppendMWikiXML = tempdoc.xml
                                               
Exit Function
ErrHandle:
    Screen.MousePointer = vbDefault
    If Len(objSystem.strErrorFunction) = 0 Then
        objSystem.strErrorModule = mc_strModule
        objSystem.strErrorFunction = "AppendMWikiXML"
    End If
    Err.Raise Err.Number
End Function

'ע��: ��ȫ·���ļ�����ץ���ļ������˺���ʹ�õ�FORѭ����֪��Ч�ʱ���InStrRev(PathName, "\")�߲���
Function ExtractFileName(PathName As String) As String
Dim x As Integer
For x = Len(PathName) To 1 Step -1
If Mid$(PathName, x, 1) = "\" Then Exit For
Next
ExtractFileName = Right$(PathName, Len(PathName) - x)
End Function


