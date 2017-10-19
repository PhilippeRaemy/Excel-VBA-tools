VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SimpleDataset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mFields() As Variant
Private mValues() As Variant ' mValues is the reference array for the SimpleDataset contents (not mFlags!)
Private mFlags() As Integer
Private Const mFlags_DELETED = 1
Private cRecord As Long

Public Function eof() As Boolean:  eof = (cRecord > UBound(mValues)): End Function
Public Function bof() As Boolean:  bof = (cRecord < LBound(mValues)): End Function
Public Sub MoveFirst(): cRecord = LBound(mValues): MoveNextUndeleted 1: End Sub
Public Sub MoveLast(): cRecord = UBound(mValues): MoveNextUndeleted -1: End Sub
Public Sub MoveNext(): cRecord = cRecord + 1: MoveNextUndeleted 1: End Sub
Public Sub MovePrevious(): cRecord = cRecord - 1: MoveNextUndeleted -1: End Sub
Public Property Get IsDeleted(Optional RowNum As Integer = -1) As Boolean
  If RowNum = -1 Then RowNum = cRecord
  If Not eof Or bof Then
    IsDeleted = (mFlags(RowNum) And mFlags_DELETED = mFlags_DELETED)
  End If
End Property
Public Property Get Fields() As Variant
  Fields = mFields
End Property
Public Property Get Count() As Integer
  Count = UBound(mValues) - LBound(mValues) + 1
End Property

Private Sub MoveNextUndeleted(Direction As Integer)
On Error GoTo Err_Proc:
GoTo Proc
Err_Proc:
  Logger.error "Error {0} in {2}: {1}", Err.Number, Err.Description, "SimpleDataset.MoveNextUndeleted"
  Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
  Resume
Proc:
  While Not (eof Or bof)
    If mFlags(cRecord) And mFlags_DELETED = mFlags_DELETED Then
      cRecord = cRecord + Direction
    Else
      Exit Sub
    End If
  Wend
End Sub
Public Sub AddField(FieldName As String)
  ReDim Preserve mFields(UBound(mFields) + 1)
  mFields(UBound(mFields)) = FieldName
End Sub
Public Sub SetFields(FieldNames As Variant) ' expecting string()
On Error GoTo Err_Proc:
GoTo Proc
Err_Proc:
  Logger.error "Error {0} in {2}: {1}", Err.Number, Err.Description, "SimpleDataset.SetFields"
  Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
  Resume
  Exit Sub
Proc:
  Class_Initialize
  mFields = FieldNames
  Exit Sub
End Sub
Public Property Get Value(FieldIdOrName As Variant) As Variant
On Error GoTo Err_Proc:
GoTo Proc
Err_Proc:
  Logger.error "Error {0} in {2}: {1}", Err.Number, Err.Description, "SimpleDataset.value"
  Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
  Resume
  Exit Sub
Proc:
  Value = mValues(cRecord)(FieldId(FieldIdOrName))
End Property
Public Property Let Value(FieldIdOrName As Variant, pValue As Variant)
On Error GoTo Err_Proc:
GoTo Proc
Err_Proc:
  Logger.error "Error {0} in {2}: {1}", Err.Number, Err.Description, "SimpleDataset.value"
  Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
  Resume
  Exit Property
Proc:
  mValues(cRecord)(FieldId(FieldIdOrName)) = pValue
End Property
Public Property Get ColumnValue(FieldIdOrName As String) As Variant
On Error GoTo Err_Proc:
GoTo Proc
Err_Proc:
  Logger.error "Error {0} in {2}: {1}", Err.Number, Err.Description, "SimpleDataset.ColumnValue"
  Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
  Resume
  Exit Sub
Proc:
  Dim c As Variant, r As Integer, fId As Integer
  fId = FieldId(FieldIdOrName)
  ReDim c(LBound(mValues) To UBound(mValues))
  For r = LBound(mValues) To UBound(mValues)
    c(r) = mValues(r)(fId)
  Next r
  ColumnValue = c
End Property
Public Property Get ColumnValuePastable(FieldIdOrName As String) As Variant
  ColumnValuePastable = Make2DArray(ColumnValue(FieldIdOrName))
End Property

Public Property Get DistinctColumnValue(FieldIdOrName As String) As Variant
On Error GoTo Err_Proc:
GoTo Proc
Err_Proc:
  Logger.error "Error {0} in {2}: {1}", Err.Number, Err.Description, "SimpleDataset.DistinctColumnValue"
  Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
  Resume
  Exit Sub
Proc:
  Dim COl As Variant: COl = ColumnValue(FieldIdOrName)
  Dim c As Variant, r As Integer, rr As Integer, found As Boolean, Value As Variant
  c = Array()
  For r = LBound(mValues) To UBound(mValues)
    found = False
    Value = COl(r)
    For rr = LBound(c) To UBound(c)
      If c(rr) = Value Then
        found = True
        Exit For
      End If
    Next rr
    If Not found Then
      ReDim Preserve c(UBound(c) + 1)
      c(UBound(c)) = Value
    End If
  Next r
  DistinctColumnValue = c
End Property
Public Function RowToDictionary() As Scripting.Dictionary
On Error GoTo Err_Proc:
GoTo Proc
Err_Proc:
  Logger.error "Error {0} in {2}: {1}", Err.Number, Err.Description, "SimpleDataset.value"
  Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
  Resume
  Exit Function
Proc:
  Dim dic As Scripting.Dictionary
  Set dic = New Scripting.Dictionary
  Dim i As Integer
  For i = LBound(mFields) To UBound(mFields)
    dic.Add mFields(i), mValues(cRecord)(i)
  Next i
  Set RowToDictionary = dic
End Function
Public Property Get FieldId(FieldIdOrName As Variant)
On Error GoTo Err_Proc:
GoTo Proc
Err_Proc:
  Logger.error "Error {0} in {2}: {1}", Err.Number, Err.Description, "SimpleDataset.FieldId"
  Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
  Resume
  Exit Sub
Proc:
  If IsNumeric(FieldIdOrName) Then
    FieldId = FieldIdOrName
    Exit Property
  End If

  For FieldId = LBound(mFields) To UBound(mFields)
    If mFields(FieldId) = FieldIdOrName Then Exit Property
  Next FieldId
  FieldId = -1
End Property
Public Sub Append(NewmValues As Variant, Optional fromIndex As Long = -1, Optional toIndex As Long = -2) ' expecting variant()
On Error GoTo Err_Proc:
GoTo Proc
Err_Proc:
  Logger.error "Error {0} in {2}: {1}", Err.Number, Err.Description, "SimpleDataset.Append"
  Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
  Resume
  Exit Sub
Proc:
  Debug.Assert IsArray(NewmValues)

  If IsMissing(fromIndex) Or IsMissing(toIndex) Or toIndex < fromIndex Then
    cRecord = UBound(mValues) + 1
    ReDim Preserve mValues(cRecord)
    ReDim Preserve mFlags(cRecord)
    mValues(cRecord) = NewmValues
  Else
    Select Case fromIndex
      Case LBound(NewmValues) To UBound(NewmValues)
        Select Case toIndex
          Case LBound(NewmValues) To UBound(NewmValues)
          Case Else: Exit Sub
        End Select
      Case Else: Exit Sub
    End Select
    Dim a As Variant, i As Integer, ia As Integer
    cRecord = UBound(mValues) + 1
    ReDim Preserve mValues(cRecord)
    ReDim Preserve mFlags(cRecord)
    a = Array()
    ReDim a(toIndex - fromIndex)
    For i = fromIndex To toIndex
      ia = i - fromIndex
      a(ia) = NewmValues(i)
    Next i
    mValues(cRecord) = a
  End If
End Sub
Public Sub ApplyToListObject(rng As listobject)
  ApplyArrayToListObject mFields, mValues, rng, True
End Sub
Public Sub ApplyToRange(rng As Range)
  ApplyArrayToRange mValues, rng, True
End Sub
Public Function Filter(FieldIdOrName As String, FiltermValues As Variant) As SimpleDataset
On Error GoTo Err_Proc:
GoTo Proc
Err_Proc:
  Logger.error "Error {0} in {2}: {1}", Err.Number, Err.Description, "SimpleDataset.Filter"
  Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
  Resume
  Exit Function
Proc:
  Dim r As Integer, f As Integer, filtervalue As Variant, toBeDeleted As Boolean
  Dim fId As Integer: fId = FieldId(FieldIdOrName)
  For r = LBound(mValues) To UBound(mValues)
    If Not IsArray(FiltermValues) Then
      If FiltermValues <> mValues(r)(fId) Then toBeDeleted = True
    Else
      toBeDeleted = True
      For Each filtervalue In FiltermValues
        If AreEquals(filtervalue, mValues(r)(fId)) Then toBeDeleted = False
      Next filtervalue
    End If
    If toBeDeleted Then mFlags(r) = mFlags(r) Or mFlags_DELETED
  Next r
  Set Filter = Me
End Function
Private Function AreEquals(a As Variant, b As Variant) As Boolean
  If IsObject(a) And IsObject(b) Then
    AreEquals = (a Is b)
  ElseIf IsDate(a) And IsDate(b) Then
    AreEquals = (CDate(a) = CDate(b))
  ElseIf IsNumeric(a) And IsNumeric(b) Then
    AreEquals = (CDbl(a) = CDbl(b))
  Else
    On Error Resume Next
    AreEquals = (CStr(a) = CStr(b))
  End If
End Function

Public Sub Delete(Optional RowNum As Integer = -1)
On Error GoTo Err_Proc:
GoTo Proc
Err_Proc:
  Logger.error "Error {0} in {2}: {1}", Err.Number, Err.Description, "SimpleDataset.Delete"
  Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
  Resume
  Exit Sub
Proc:
  If RowNum = -1 Then RowNum = cRecord
  mFlags(RowNum) = mFlags(RowNum) Or mFlags_DELETED
  End Sub
Public Sub Undelete(Optional RowNum As Integer = -1)
On Error GoTo Err_Proc:
GoTo Proc
Err_Proc:
  Logger.error "Error {0} in {2}: {1}", Err.Number, Err.Description, "SimpleDataset.Undelete"
  Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
  Resume
  Exit Sub
Proc:
  If RowNum = -1 Then RowNum = cRecord
  mFlags(RowNum) = mFlags(RowNum) Xor mFlags_DELETED:
End Sub
Public Sub UndeleteAll()
On Error GoTo Err_Proc:
GoTo Proc
Err_Proc:
  Logger.error "Error {0} in {2}: {1}", Err.Number, Err.Description, "SimpleDataset.UndeleteAll"
  Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
  Resume
  Exit Sub
Proc:
  Dim r As Integer
  For r = LBound(mValues) To UBound(mValues)
    Undelete r
  Next r
End Sub
Public Function Pack() As SimpleDataset
On Error GoTo Err_Proc:
GoTo Proc
Err_Proc:
  Logger.error "Error {0} in {2}: {1}", Err.Number, Err.Description, "SimpleDataset.Pack"
  Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
  Resume
  Exit Function
Proc:
  Dim packed As SimpleDataset
  Set packed = New SimpleDataset
  packed.SetFields mFields
  Me.AppendTo packed, IncludeDeleted:=False
  Set Pack = packed
End Function
Friend Sub AppendTo(other As SimpleDataset, Optional IncludeDeleted As Boolean = True)
On Error GoTo Err_Proc:
GoTo Proc
Err_Proc:
  Logger.error "Error {0} in {2}: {1}", Err.Number, Err.Description, "SimpleDataset.AppendTo"
  Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
  Resume
  Exit Sub
Proc:
  Dim o As Integer
  For o = LBound(mValues) To UBound(mValues)
    If IncludeDeleted Or Not IsDeleted(o) Then
      other.Append mValues(o)
      If IsDeleted(o) Then other.Delete
    End If
  Next o
End Sub
Public Function Clone() As SimpleDataset
On Error GoTo Err_Proc:
GoTo Proc
Err_Proc:
  Logger.error "Error {0} in {2}: {1}", Err.Number, Err.Description, "SimpleDataset.Clone"
  Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
  Resume
  Exit Function
Proc:
  Dim o As Integer
  Dim cloned As SimpleDataset
  Set cloned = New SimpleDataset
  cloned.SetFields mFields
  Me.AppendTo cloned
  Set Clone = cloned
End Function

Public Function Concat(other As SimpleDataset) As SimpleDataset
On Error GoTo Err_Proc:
GoTo Proc
Err_Proc:
  Logger.error "Error {0} in {2}: {1}", Err.Number, Err.Description, "SimpleDataset.Concat"
  Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
  Resume
  Exit Function
Proc:
  ' assuming same mFields...
  Dim cloned As SimpleDataset
  Set cloned = Me.Clone
  other.AppendTo cloned
  Set Concat = cloned
End Function

Public Sub FromListObject(lo As listobject)
On Error GoTo Err_Proc:
GoTo Proc
Err_Proc:
  Logger.error "Error {0} in {2}: {1}", Err.Number, Err.Description, "SimpleDataset.FromListObject"
  Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
  Resume
  Exit Sub
Proc:
  Class_Initialize
  mFields = MakeArrayOfArray(lo.HeaderRowRange.Value)(1)
  If Not lo.DataBodyRange Is Nothing Then
    mValues = MakeArrayOfArray(lo.DataBodyRange.Value)
    ReDim mFlags(LBound(mValues) To UBound(mValues))
  End If
End Sub
Public Sub FromServiceXML(xml As MSXML2.DOMDocument60)
  Dim xfields As MSXML2.IXMLDOMNodeList
  Dim xfield As MSXML2.IXMLDOMNode
  Dim xrows As MSXML2.IXMLDOMNodeList
  Dim xrow As MSXML2.IXMLDOMNode
  Dim fi As Integer, Row As Integer
  Dim v As Variant
  Set xfields = xml.SelectNodes("//a:Fields/b:string")
  ReDim Preserve mFields(xfields.Length - 1)
  fi = 0
  For Each xfield In xfields
    mFields(fi) = xfield.Text
    fi = fi + 1
  Next xfield
  Set xrows = xml.SelectNodes("//a:Values/b:ArrayOfstring")
  ReDim Preserve mValues(xrows.Length - 1)
  Row = 0
  For Each xrow In xrows
    fi = 0
    v = Array(): ReDim v(UBound(mFields))
    Set xfields = xrow.SelectNodes("b:string")
    For Each xfield In xfields
      v(fi) = xfield.Text
      fi = fi + 1
    Next xfield
    mValues(Row) = v
    Row = Row + 1
  Next xrow
End Sub

Public Function Serialize(rootNodeName As String, ParamArray rootNodeAttributes() As Variant) As MSXML2.DOMDocument60
  Dim doc As MSXML2.DOMDocument60, root As MSXML2.IXMLDOMNode
  Dim FieldsNode As MSXML2.IXMLDOMNode, ValuesNode As MSXML2.IXMLDOMNode, RowNode As MSXML2.IXMLDOMNode, ChildNode As MSXML2.IXMLDOMNode
  Dim fi As Variant, Row As Variant
  Set doc = New MSXML2.DOMDocument60
  Set root = XMLHelpers.AppendChildElement(doc, rootNodeName)
  Dim rootNodeAttribute As Variant
  For Each rootNodeAttribute In rootNodeAttributes
    rootNodeAttribute = Split(rootNodeAttribute, "=")
    XMLHelpers.SetAttribute root, CStr(rootNodeAttribute(0)), CStr(rootNodeAttribute(1))
  Next rootNodeAttribute
  Set FieldsNode = XMLHelpers.AppendChildElement(root, "a:Fields")
  Set ValuesNode = XMLHelpers.AppendChildElement(root, "a:Values")
  XMLHelpers.SetAttribute FieldsNode, "xmlns:b", "http://schemas.microsoft.com/2003/10/Serialization/Arrays"
  XMLHelpers.SetAttribute ValuesNode, "xmlns:b", "http://schemas.microsoft.com/2003/10/Serialization/Arrays"
  
  For Each fi In Me.Fields
    Set ChildNode = XMLHelpers.AppendChildElement(FieldsNode, "b:string", CStr(fi))
  Next fi
  For Each Row In mValues
    Set ChildNode = XMLHelpers.AppendChildElement(ValuesNode, "b:ArrayOfstring")
    For Each fi In Row
      XMLHelpers.AppendChildElement ChildNode, "b:string", CStr(fi)
    Next fi
  Next Row
  Set Serialize = doc
End Function

Private Sub Class_Initialize()
  mValues = Array()
  mFields = Array()
  ReDim mFlags(0)
  cRecord = -1
End Sub
' make variant(height, width) from variant(height)(width)
Private Function Make2DArray(i As Variant, Optional Width As Integer = -1) As Variant
On Error GoTo Err_Proc:
GoTo Proc
Err_Proc:
  Logger.error "Error {0} in {2}: {1}", Err.Number, Err.Description, "SimpleDataset.Make2DArray"
  Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
  Resume
  Exit Function
Proc:
Dim r As Integer, c As Integer, a() As Variant
  If Not IsArray(i) Then Exit Function
  If UBound(i) = -1 Then Exit Function
  If Not IsArray(i(LBound(i))) Then
    Width = 0
  ElseIf Width = -1 Then
    Width = UBound(i(LBound(i)))
  End If
  ReDim a(LBound(i) To UBound(i), Width)
  For r = LBound(i) To UBound(i)
    For c = 0 To Width
      If IsArray(i(r)) Then
        If c <= UBound(i(r)) Then
          a(r, c) = i(r)(c)
        End If
      Else
        a(r, c) = i(r)
      End If
    Next c
  Next r
  Make2DArray = a
End Function
Private Function Dimensions(a As Variant) As Integer
  Dim D As Integer
  Dimensions = 0
  On Error GoTo EndFunction
  While True: D = UBound(a, Dimensions + 1): Dimensions = Dimensions + 1: Wend
EndFunction:
End Function
Private Function MakeArrayOfArray(a As Variant) As Variant
On Error GoTo Err_Proc:
GoTo Proc
Err_Proc:
  Logger.error "Error {0} in {2}: {1}", Err.Number, Err.Description, "SimpleDataset.MakeArrayOfArray"
  Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
  Resume
  Exit Function
Proc:

Dim r As Variant, c As Variant, i As Integer, j As Integer
  r = Array()
  Select Case Dimensions(a)
    Case 0: ' ignore
    Case 1:
      ReDim r(LBound(a) To UBound(a))
      For i = LBound(a) To UBound(a)
        r(i) = Array(a(i))
      Next i
    Case Is >= 2:
      ReDim r(LBound(a, 1) To UBound(a, 1))
      For i = LBound(a, 1) To UBound(a, 1)
        c = Array()
        ReDim c(LBound(a, 2) To UBound(a, 2))
        For j = LBound(a, 2) To UBound(a, 2)
          c(j) = a(i, j)
        Next j
        r(i) = c
      Next i
  End Select
  MakeArrayOfArray = r
End Function

Private Sub ApplyArrayToRange(i As Variant, rng As Range, Optional overflow As Boolean = False)
On Error GoTo Err_Proc:
GoTo Proc
Err_Proc:
  Logger.error "Error {0} in {2}: {1}", Err.Number, Err.Description, "SimpleDataset.ApplyArrayToRange"
  Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
  Resume
Proc:
  If Not rng Is Nothing Then rng.Clear
  Dim a As Variant
  a = Make2DArray(i)
  If IsEmpty(a) Then Exit Sub
  Dim targetRng As Range
  If Not overflow Then
    Set targetRng = rng
  Else
    Set targetRng = rng.parent.Range(rng.Cells(1, 1), rng.Cells(UBound(a, 1) + 1, UBound(a, 2) + 1))
  End If
  targetRng.Value = a
End Sub

Private Sub ApplyArrayToListObject(hdr As Variant, data As Variant, lo As listobject, Optional resize As Boolean = False)
On Error GoTo Err_Proc:
GoTo Proc
Err_Proc:
  Logger.error "Error {0} in {2}: {1}", Err.Number, Err.Description, "SimpleDataset.ApplyArrayToListObject"
  Err.Raise Err.Number, Err.source, Err.Description, Err.HelpFile, Err.HelpContext
  Resume
  Exit Sub
Proc:
  If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Cells.Clear
  If resize Then
      RangeHelpers.setListObjectSize lo, UBound(hdr) + 1, UBound(data) + 1
  End If
  If UBound(data) < 0 Then Exit Sub
  
  ApplyArrayToRange data, IIf(lo.DataBodyRange Is Nothing, lo.InsertRowRange, lo.DataBodyRange), resize
  lo.HeaderRowRange.Value = hdr
End Sub


