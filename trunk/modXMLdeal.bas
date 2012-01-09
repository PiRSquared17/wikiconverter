Attribute VB_Name = "modXMLdeal"
Public Sub AppendWikiXML(strRootFolder As String, strFile As String)

Dim m_AppPath As String
Dim xml_document As DOMDocument
Dim values_node As IXMLDOMNode
Dim values_node2 As IXMLDOMNode
Dim Root_Node As IXMLDOMNode
Dim n As Integer

    '获取程序路径
    m_AppPath = App.Path
    If Right$(m_AppPath, 1) <> "\" Then m_AppPath = m_AppPath & "\"
    '加载文档
    Set xml_document = New DOMDocument
    xml_document.Load m_AppPath & "export2.xml"
    '如果不包含任何元素，则退出
    If xml_document.documentElement Is Nothing Then
        Exit Sub
    End If
    'Root_Node置为根节点
    Set Root_Node = xml_document.selectSingleNode("mediawiki")
    '获取Total节点的Text
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
    
    '添加新的节点
    
    'N = CInt(GetNodeValue(Root_Node, "Total", "?")) + 1
    '添加新元素
    CreateNode 4, Root_Node, "page", ""
    Set values_node = Root_Node.selectSingleNode("Page")
    CreateNode 4, values_node, "Name", "姓名"
    CreateNode 4, values_node, "FileName", "2006"
    CreateNode 4, values_node, "Del", "0"
    '修改Total值
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
'Desctiption:      创建一份MEDIAWIKI格式的Xml
'                       格式如下:符合
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
'                              <comment>以"hello world"为内容创建页面</comment>
'                              <text xml:space="preserve" bytes="11">hello world</text>
'                            </revision>
'                          </page>
'                        </mediawiki>
'
'Parameters:       strPagePath
'
'Return Value:      返回对应的xml文檔
'*************************************************************************
'DATE               NAME                    DESCRIPTION
'-------------------------------------------------------------------------
'2012-01-07      QingHui Chen               Function create
'*************************************************************************
Public Function AppendMWikiXML(ByVal strPagePath As String) As String
                                               
On Error GoTo ErrHandle:

    Dim tempdoc As MSXML2.DOMDocument               '定义的xml文件变量
    Dim EL_curElement As MSXML2.IXMLDOMElement      '定义根节点
    Dim RootNode As MSXML2.IXMLDOMNode
    Dim l_IXMPI As IXMLDOMProcessingInstruction
    Dim l_strxml As String                          '生成的XML文檔String
   
   
    Dim node_Parent As MSXML2.IXMLDOMNode '父节点对像
    Dim node_son As MSXML2.IXMLDOMNode '子节点对像
    Dim node_grandson As MSXML2.IXMLDOMNode '孙节点对像
    Dim node_greatgrandson As MSXML2.IXMLDOMNode '玄孙节点对像
    Dim tempAttribute As IXMLDOMElement
    
   
    Dim i As Integer                                        '临时变量
   
    If Len(strPagePath) = 0 Then
        AppendMWikiXML = ""
        Exit Function
    End If
   
    Set tempdoc = New MSXML2.DOMDocument
    Set EL_curElement = tempdoc.createElement("mediawiki")     '这里mediawiki作为根接点
   
    Set tempdoc.documentElement = EL_curElement
    
    
    
' 递归搜索文件直到无结果
   i = 0
   frmMain.List2.Clear
      Do Until i = 1000   '本次待处理的为786个TXT文件
        If i < 0 Then Exit Do
        'I = I + 1
        strFilePath = FindOnefile$(strPagePath, "*.txt")
        If strFilePath = "" Then Exit Do
        
        If Right$(strFilePath, 4) = ".txt" Then
            strFileTempPath = Replace$(strFilePath, ".txt", ".bin")
            Name strFilePath As strFileTempPath
            strFilePath = strFileTempPath
            frmMain.List2.AddItem "文件 | " & ExtractFileName("" & strFilePath)
                 Set node_Parent = tempdoc.CreateNode(MSXML2.NODE_ELEMENT, "page", "")  '创建page节点
                
                 EL_curElement.appendChild node_Parent   '在mediawiki下加入 page 节点
                 
                 Set node_son = tempdoc.CreateNode(MSXML2.NODE_ELEMENT, "title", "")  '创建 title 节点
                 node_Parent.appendChild node_son   '在 page 下加入 title 节点
                 node_son.Text = ExtractFileName("" & strFilePath)   '设定 title 节点值
                 node_son.Text = Replace(node_son.Text, ".bin", "")

                 
                 Set node_son = tempdoc.CreateNode(MSXML2.NODE_ELEMENT, "id", "")  '创建 id 节点
                 node_Parent.appendChild node_son   '在 page 下加入 id 节点
                 'If strFilePath <> "" And strFilePath <> "wikiconverter.exe" Then i = i + 1
                 i = i + 1
                 node_son.Text = i   '设定 id 节点值
                 
                 Set node_son = tempdoc.CreateNode(MSXML2.NODE_ELEMENT, "revision", "")  '创建 revision 节点
                 node_Parent.appendChild node_son   '在 page 下加入 revision 节点
                 
                 Set node_grandson = tempdoc.CreateNode(MSXML2.NODE_ELEMENT, "id", "")  '创建 revision id 节点
                 node_son.appendChild node_grandson   '在 revision 下加入 id 节点
                 node_grandson.Text = "1"     '设定 revision id 节点值
                 Set node_grandson = tempdoc.CreateNode(MSXML2.NODE_ELEMENT, "timestamp", "")  '创建 revision timestamp 节点
                 node_son.appendChild node_grandson   '在 revision 下加入 timestamp 节点
                 node_grandson.Text = "2012-01-08T08:08:28Z"  '设定 revision timestamp 节点值
                 Set node_grandson = tempdoc.CreateNode(MSXML2.NODE_ELEMENT, "contributor", "")  '创建 revision contributor 节点
                 node_son.appendChild node_grandson   '在 revision 下加入 contributor 节点
                 
                 Set node_greatgrandson = tempdoc.CreateNode(MSXML2.NODE_ELEMENT, "username", "")  '创建 revision contributor username 节点
                 node_grandson.appendChild node_greatgrandson   '在 contributor 下加入  revision contributor username  节点
                 node_greatgrandson.Text = "Dragon"  '设定 revision contributor username 节点值
                 Set node_greatgrandson = tempdoc.CreateNode(MSXML2.NODE_ELEMENT, "id", "")  '创建 revision contributor id 节点
                 node_grandson.appendChild node_greatgrandson   '在 contributor 下加入  revision contributor id  节点
                 node_greatgrandson.Text = "2"  '设定 revision contributor id 节点值
                 
                 Set node_grandson = tempdoc.CreateNode(MSXML2.NODE_ELEMENT, "comment", "")  '创建 revision comment 节点
                 node_son.appendChild node_grandson   '在 revision 下加入 comment 节点
                 node_grandson.Text = "from Dokuwiki"  '设定 revision contributor comment 节点值
                 
                 Set node_grandson = tempdoc.CreateNode(MSXML2.NODE_ELEMENT, "text", "")  '创建 revision text 节点
                 node_son.appendChild node_grandson   '在 revision 下加入 text 节点
                 
                 node_grandson.Text = File_get_contents("" & strFilePath, "UTF-8")  '设定 revision text 节点值

                 '<<---------创建 revision text里的 Attribute 属性及设定值--------------
                 
                 Set tempAttribute = node_grandson
                 tempAttribute.setAttribute "xml:space", "preserve"
                 tempAttribute.setAttribute "bytes", txtfilesize

        End If
        
      DoEvents
   Loop
   
   
    
' 把临时改为的bin扩展名文件恢复原来的TXT扩展名
   i = 0
   Do Until i = 1000   '本次待处理的为786个TXT文件
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
   
    '测试试一下有无成功生成
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

'注释: 从全路径文件名中抓出文件名，此函数使用的FOR循环不知道效率比起InStrRev(PathName, "\")高不？
Function ExtractFileName(PathName As String) As String
Dim x As Integer
For x = Len(PathName) To 1 Step -1
If Mid$(PathName, x, 1) = "\" Then Exit For
Next
ExtractFileName = Right$(PathName, Len(PathName) - x)
End Function


