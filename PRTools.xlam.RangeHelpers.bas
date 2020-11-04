Attribute VB_Name = "RangeHelpers"
Option Explicit

Public Enum EnumRangeRelation
    Disjointed
    Including
    Included
    Overlapping
End Enum

Sub setListObjectWidth(table As listobject, Width As Integer)
    setListObjectSize table, Width
End Sub
Sub setListObjectHeight(table As listobject, height As Integer)
    setListObjectSize table, , height
End Sub

Sub setListObjectSize(table As listobject, Optional Width As Integer = -1, Optional height As Integer = -1)
    Dim r As Range: Set r = table.Range
    Select Case Width
        Case Is < 0: Width = r.columns.Count
        Case Is < 1: Width = 1
    End Select
    Select Case height
        Case Is < 0: height = r.Rows.Count
        Case Is < 1: height = 1
    End Select
    table.resize Range(r.Cells(1, 1), r.Cells(height + 1, Width))
End Sub

Function SplitAll(a As Variant, delimiter As String) As Variant
    Dim r As Variant, s As Variant, E As Variant, i As Integer, ii As Integer
    If Not IsArray(a) Then
        SplitAll = Split(CStr(a), delimiter)
        Exit Function
    End If
    r = Array()
    For Each E In a
        s = Split(CStr(E), delimiter)
        ReDim Preserve r(UBound(r) + UBound(s) + 1)
        For ii = 0 To UBound(s)
            r(i) = s(ii)
            i = i + 1
        Next ii
    Next E
    SplitAll = r
End Function
Sub AlignButtons()
Const Width = 80
Const height = 40
Const space = 3
    Dim shp As Shape
    Dim prevshp As Shape
    For Each shp In ActiveSheet.Shapes
        If prevshp Is Nothing Then
            shp.Left = shp.TopLeftCell.Offset(0, 0).Left + space
            shp.Width = Width
            shp.height = height
            Set prevshp = shp
        Else
            shp.Left = prevshp.Left + prevshp.Width + space
            shp.Top = prevshp.Top
            shp.Width = Width
            shp.height = height
            Set prevshp = shp
        End If
        shp.Placement = xlFreeFloating
    Next shp
End Sub

Public Function ColumnName(columnNumber As Integer) As String
    Select Case columnNumber
        Case Is <= 26: ColumnName = Chr(64 + columnNumber)
        Case Is <= 702: ColumnName = Chr(65 + Int((columnNumber - 27) / 26)) & Chr(65 + ((columnNumber - 1) Mod 26))
        Case Is <= 18278: ColumnName = Chr(65 + Int((columnNumber - 703) / 676)) & Chr(65 + Int((columnNumber - 703) / 26) Mod 26) & Chr(65 + ((columnNumber - 1) Mod 26))
    End Select
End Function
Private Sub testcolumnsname()
    Dim c As Integer, proofname As String, calcName As String
    For c = 1 To 16384
        proofname = Mid(ActiveSheet.Cells(1, c).Address, 2, InStr(2, ActiveSheet.Cells(1, c).Address, "$") - 2)
        calcName = ColumnName(c)
        If Not calcName = proofname Then
            Debug.Print c, calcName, proofname
            Stop
        End If
    Next c
    Debug.Print "Done!"
End Sub
Public Function GetColumnsOrdinalDictionary(rng As Range, ParamArray ColumnNames() As Variant)
    Dim colName As Variant
    Dim dic As scripting.Dictionary: Set dic = New scripting.Dictionary
    
    For Each colName In ColumnNames
        dic(colName) = Application.WorksheetFunction.Match(colName, rng, False)
    Next colName
    Set GetColumnsOrdinalDictionary = dic
End Function

Public Function ToStringArray(rng As Range) As Variant
    Dim cell As Range, i As Integer, a() As String
    ReDim a(rng.Rows.Count * rng.columns.Count - 1)
    For Each cell In rng
        a(i) = CStr(cell.Value)
        i = i + 1
    Next cell
    ToStringArray = a
End Function

Public Function RangeRelation(r1 As Range, r2 As Range) As EnumRangeRelation
Dim hRelation As String
Dim vRelation As String
    hRelation = IntervalRelation(r1.column, r1.column + r1.columns.Count, r2.column, r2.column + r2.columns.Count)
    vRelation = IntervalRelation(r1.Row, r1.Row + r1.Rows.Count, r2.Row, r2.Row + r2.Rows.Count)
    If hRelation = vRelation Then
        RangeRelation = vRelation
    Else
        RangeRelation = Disjointed
    End If
End Function

Private Function IntervalRelation(x1 As Long, x2 As Long, y1 As Long, y2 As Long) As EnumRangeRelation
    If x2 < y1 Then
        IntervalRelation = Disjointed
    ElseIf y2 < x1 Then
        IntervalRelation = Disjointed
    ElseIf x1 <= y1 And x2 >= y2 Then
        IntervalRelation = Including
    ElseIf y1 <= x1 And y2 >= x2 Then
        IntervalRelation = Included
    Else
        IntervalRelation = Overlapping
    End If
End Function

Public Function CountDistinct(r As Range) As Integer
Dim a As Variant, d As Dictionary, v As Variant
    a = r.Value
    Set d = New Dictionary
    For Each v In a
        If Not d.Exists(v) Then
            CountDistinct = CountDistinct + 1
            d.Add v, Empty
        End If
    Next v
End Function
