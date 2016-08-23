Attribute VB_Name = "XMLHelpers"
'## pragma PRTools export
Option Explicit

Public Function DumpNode(node As IXMLDOMNode, Optional indent As String = "") As String
  Dim child As IXMLDOMNode
  Dim attr As IXMLDOMAttribute
  Dim nodeText As String
  Dim ChildDumps As String
  Dim children As Variant, c As Long
  
  
  DumpNode = indent & "<" & node.BaseName
  If Not node.Attributes Is Nothing Then
    For Each attr In node.Attributes
      DumpNode = DumpNode & " " & attr.BaseName & " = """ & attr.value & """"
    Next attr
  End If
  DumpNode = DumpNode & ">"
  
  children = Array()
  ReDim children(node.ChildNodes.Length - 1)
  For c = 0 To node.ChildNodes.Length - 1
    children(c) = Array(node.ChildNodes(c).xml, node.ChildNodes(c))
  Next c
  QuickSort children, LBound(children), UBound(children)
  For c = LBound(children) To UBound(children)
    Set child = children(c)(1)
    If child.BaseName = "" Then
      nodeText = child.Text
    Else
      ChildDumps = ChildDumps & DumpNode(child, indent & "  ")
    End If
  Next c
  If ChildDumps <> "" And nodeText <> "" Then
    DumpNode = DumpNode & vbCrLf & indent & node.Text & vbCrLf & ChildDumps & vbCrLf
  ElseIf ChildDumps <> "" Then
    DumpNode = DumpNode & vbCrLf & ChildDumps & indent
  Else
    DumpNode = DumpNode & node.Text
  End If
  DumpNode = DumpNode & "</" & node.BaseName & ">" & vbCrLf
End Function
Private Function LooseLoadXml(s As String) As DOMDocument
Dim node As New DOMDocument
  If Not node.LoadXML(s) Then
    If Not node.LoadXML(Replace(s, ">", "/>")) Then
      Err.Raise vbObject, "LooseLoadXML", "Invalid XML stream"
    End If
  End If
  Set LooseLoadXml = node
End Function
Public Function GetXMLTagName(s As String)
Dim node As DOMDocument
  Set node = LooseLoadXml(s)
  If Not node Is Nothing Then
    GetXMLTagName = node.FirstChild.NodeName
  End If
End Function

Public Function GetXMLTagText(s As String)
Dim node As DOMDocument
  Set node = LooseLoadXml(s)
  If Not node Is Nothing Then
    GetXMLTagText = node.FirstChild.Text
  End If
End Function

Public Sub CreateChildren(root As MSXML2.DOMDocument, parent As MSXML2.IXMLDOMNode, nodes As MSXML2.DOMDocument)
Dim child As IXMLDOMNode
  For Each child In nodes.ChildNodes
    parent.appendChild nodes.FirstChild
  Next child
End Sub
Public Function UpSertNode(parent As MSXML2.IXMLDOMNode, NodeName As String, value As String) As IXMLDOMNode
  Dim node As IXMLDOMNode
  Dim root As New MSXML2.DOMDocument
  For Each node In parent.ChildNodes
    If node.NodeName = NodeName Then
      Set UpSertNode = node
      Exit For
    End If
  Next node
  If UpSertNode Is Nothing Then
    Set UpSertNode = OwnerDocument(parent).CreateElement(NodeName)
    parent.appendChild UpSertNode
  End If
  UpSertNode.Text = value
End Function

Public Sub setNodeText(parent As IXMLDOMNode, xPathSelector As String, value As String)
  Dim node As IXMLDOMNode
  For Each node In parent.SelectNodes(xPathSelector)
    node.Text = value
  Next node
End Sub
Private Function OwnerDocument(node As IXMLDOMNode) As DOMDocument
  If TypeName(node) = "DOMDocument" Then
    Set OwnerDocument = node
  Else
    Set OwnerDocument = node.OwnerDocument
  End If
End Function

Public Function AppendChildElement(parent As IXMLDOMNode, ElementName As String, Optional ElementText As String) As IXMLDOMNode
  Dim doc As MSXML2.DOMDocument
  Set AppendChildElement = parent.appendChild(OwnerDocument(parent).CreateElement(ElementName))
  If Not isEmptyOrBlank(ElementText) Then
    AppendChildElement.Text = ElementText
  End If
End Function

Public Function SetAttribute(node As IXMLDOMNode, AttributeName As String, AttributeValue As String) As IXMLDOMAttribute
  Set SetAttribute = OwnerDocument(node).createAttribute(AttributeName)
  SetAttribute.value = AttributeValue
  node.Attributes.setNamedItem SetAttribute
End Function

Public Function GetNodeAttributeText(node As IXMLDOMNode, AttributeName As String, Optional default As String) As String
  GetNodeAttributeText = default
  On Error Resume Next
  GetNodeAttributeText = node.Attributes.getNamedItem(AttributeName).Text
  If Trim(GetNodeAttributeText) = "" Then GetNodeAttributeText = default
End Function
Public Function GetNodeAttributeDbl(node As IXMLDOMNode, AttributeName As String, Optional default As Double) As Double
  GetNodeAttributeDbl = default
  On Error Resume Next
  GetNodeAttributeDbl = CDbl(node.Attributes.getNamedItem(AttributeName).Text)
End Function
Public Function GetNodeAttributeDate(node As IXMLDOMNode, AttributeName As String, Optional default As Date) As Date
  GetNodeAttributeDate = default
  On Error Resume Next
  GetNodeAttributeDate = cIsoDate(node.Attributes.getNamedItem(AttributeName).Text)
End Function
Public Function GetNodeChildText(node As IXMLDOMNode, xPathSelector As String, Optional default As String) As String
  GetNodeChildText = default
  On Error Resume Next
  GetNodeChildText = node.SelectSingleNode(xPathSelector).Text
  If Trim(GetNodeChildText) = "" Then GetNodeChildText = default
End Function
Public Function GetNodeChildDbl(node As IXMLDOMNode, xPathSelector As String, Optional default As Double) As Double
  GetNodeChildDbl = default
  On Error Resume Next
  GetNodeChildDbl = CDbl(node.SelectSingleNode(xPathSelector).Text)
End Function
Public Function GetNodeChildDate(node As IXMLDOMNode, xPathSelector As String, Optional default As Date) As Date
  GetNodeChildDate = default
  On Error Resume Next
  GetNodeChildDate = cIsoDate(node.SelectSingleNode(xPathSelector).Text)
End Function
Private Function cIsoDate(tIsoDate As String) As Date
  On Error Resume Next
  cIsoDate = CDate(Mid(tIsoDate, 1, 10) & " " & Mid(tIsoDate, 12))
End Function


Public Sub testXPath()
Dim doc As New MSXML2.DOMDocument
Dim xml As String
Dim node As MSXML2.IXMLDOMNode
xml = "<root>"
xml = xml & "<node><key>a</key><value>1</value></node>"
xml = xml & "<node><key>b</key><value>2</value></node>"
xml = xml & "<node><key>c</key><value>3</value></node>"
xml = xml & "<subnodes>"
xml = xml & "<node><key>a</key><value>s1</value></node>"
xml = xml & "<node><key>b</key><value>s2</value></node>"
xml = xml & "<node><key>c</key><value>s3</value></node>"
xml = xml & "</subnodes>"
xml = xml & "</root>"

doc.LoadXML xml
For Each node In doc.SelectNodes("/root/node[key='c']"): Debug.Print node.xml: Next node
For Each node In doc.SelectNodes("/root/node[key='c']"): Debug.Print node.SelectSingleNode("value").Text: Next node
For Each node In doc.SelectNodes("//node[key='c']"): Debug.Print node.SelectSingleNode(".//value").Text: Next node

End Sub

Public Sub QuickSort(arr, lo As Long, Hi As Long)
  Dim varPivot As Variant
  Dim varTmp As Variant
  Dim tmpLow As Long
  Dim tmpHi As Long
  tmpLow = lo
  tmpHi = Hi
  varPivot = arr((lo + Hi) \ 2)(0)
  Do While tmpLow <= tmpHi
    Do While arr(tmpLow)(0) < varPivot And tmpLow < Hi
      tmpLow = tmpLow + 1
    Loop
    Do While varPivot < arr(tmpHi)(0) And tmpHi > lo
      tmpHi = tmpHi - 1
    Loop
    If tmpLow <= tmpHi Then
      varTmp = arr(tmpLow)
      arr(tmpLow) = arr(tmpHi)
      arr(tmpHi) = varTmp
      tmpLow = tmpLow + 1
      tmpHi = tmpHi - 1
    End If
  Loop
  If lo < tmpHi Then QuickSort arr, lo, tmpHi
  If tmpLow < Hi Then QuickSort arr, tmpLow, Hi
End Sub

