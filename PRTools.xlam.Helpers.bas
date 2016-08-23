Attribute VB_Name = "Helpers"
' ####################
' \\GVA0MS01\RAEMYP\Excel\Copy of PRTools.xlsm.Helpers.bas
' ####################
Sub CreateODBCQuery(sql As String, tag As String)
  With sh.ListObjects.Add _
      (SourceType:=0 _
      , source:=Array( _
          "OLEDB;Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;Data Source=kstlon0db003;Use Procedure for Prepare=1;" _
        , "Auto Translate=True;Packet Size=4096;Use Encryption for Data=False;Tag with column collation when possible=" _
        , "False;Initial Catalog=TimeSeries") _
      , Destination:=Range("$A$2") _
      ).QueryTable
      .CommandType = xlCmdSql
      .CommandText = Array(sql)
      .RowNumbers = False
      .FillAdjacentFormulas = False
      .PreserveFormatting = True
      .RefreshOnFileOpen = False
      .BackgroundQuery = True
      .RefreshStyle = xlInsertDeleteCells
      .SavePassword = False
      .SaveData = True
      .AdjustColumnWidth = True
      .RefreshPeriod = 0
      .PreserveColumnInfo = True
      .listobject.DisplayName = "Table_ExternalData_" & tag
      .Refresh BackgroundQuery:=False
    End With
End Sub


Sub ListQueries()
  Dim qry As WorkbookConnection
  For Each qry In ActiveWorkbook.Connections
    Debug.Print qry.Name,
    On Error Resume Next
    Debug.Print qry.OLEDBConnection.CommandText;
    Debug.Print
  Next qry
End Sub

Function MakeRange(ParamArray ranges() As Variant) As Range
Dim rng As Range, i As Integer, addr As String
    For i = 0 To UBound(ranges)
        Set rng = ranges(i)
        addr = addr & rng.Address & ","
    Next i
    Set MakeRange = Range(Left(addr, Len(addr) - 1))
End Function

Function MakeInsert(TableName As String, Columns As Variant, Values As Variant, Optional union As String) As String
Dim cell As Range, ValueExpression As String, i As Integer, colNames() As String
    If Not Columns.Cells.Count = Values.Cells.Count Then
        MakeInsert = "columns and values do not match"
        Exit Function
    End If
    If union = "" Then
        MakeInsert = "INSERT INTO [" & TableName & "]("
    Else
        ReDim colNames(Columns.Cells.Count - 1)
    End If
    For Each cell In Columns.Cells
        If union = "" Then
            MakeInsert = MakeInsert & "[" & cell.value & "], "
        Else
            colNames(i) = "[" & cell.value & "]="
            i = i + 1
        End If
    Next cell
    i = 0
    For Each cell In Values.Cells
        If union <> "" Then
            ValueExpression = ValueExpression & colNames(i)
            i = i + 1
        End If
        ValueExpression = ValueExpression & MakeSqlLiteral(cell.value) & ","
    Next cell
    If union = "" Then
        MakeInsert = Left(MakeInsert, Len(MakeInsert) - 2) & ") SELECT " & Left(ValueExpression, Len(ValueExpression) - 1)
    Else
        MakeInsert = union & "SELECT " & Left(ValueExpression, Len(ValueExpression) - 1)
    End If
End Function
Function MakeSelectExpression(Values As Range) As String
Dim cell As Range
    For Each cell In Values.Cells
        MakeSelectExpression = MakeSelectExpression & MakeSqlLiteral(cell.value) & ","
    Next cell
    MakeSelectExpression = Left(MakeSelectExpression, Len(MakeSelectExpression) - 1)
End Function
Function MakeSqlLiteral(v As Variant) As String
  If IsNumeric(v) Then
    MakeSqlLiteral = "'" & CStr(v) & "'"
  ElseIf CStr(v) Like "(SELECT *)" Then
    MakeSqlLiteral = CStr(v)
  Else
    MakeSqlLiteral = "'" & Replace(CStr(v), "'", "''") & "'"
  End If
End Function

Function Join(Sep As String, rng As Range) As String
Dim cell As Range
  For Each cell In rng
    If Join <> "" Then Join = Join & Sep
    Join = Join & cell.value
  Next cell
End Function


Sub BubbleSort(arr)
  Dim swap As Variant
  Dim i As Long
  Dim j As Long
  Dim lngMin As Long
  Dim lngMax As Long
  lngMin = LBound(arr)
  lngMax = UBound(arr)
  For i = lngMin To lngMax - 1
    For j = i + 1 To lngMax
      If arr(i)(0) > arr(j)(0) Then
        swap = arr(i)
        arr(i) = arr(j)
        arr(j) = swap
      End If
    Next j
  Next i
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
