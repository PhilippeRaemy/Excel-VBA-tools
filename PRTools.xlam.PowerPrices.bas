Attribute VB_Name = "PowerPrices"
Option Explicit

Sub Create_MDE_FILE()

Dim lastCell As Range

Dim Dates As Variant
Dim Series As Variant
Dim Values As Variant
Dim ValuesRange As Range
Dim SeriesNames() As String
Dim SeriesName As Variant
Dim Observations() As String
Dim c As Integer, r As Integer
Dim fso As New FileSystemObject
Dim ts As TextStream
Const OUTFILENAME = "c:\temp\Gundies_CXL_RECOVERY_20150227_000000.csv"
    
    Set ts = fso.OpenTextFile(OUTFILENAME, ForAppending, True)
    Set lastCell = ActiveCell.SpecialCells(xlLastCell)
    Dates = ActiveSheet.Range("f1", ActiveSheet.Cells(1, lastCell.column)).Value
    Series = ActiveSheet.Range("e2", ActiveSheet.Cells(lastCell.Row, 5)).Value
    Values = ActiveSheet.Range("f2", lastCell).Value
    ReDim SeriesNames(LBound(Values, 1) To UBound(Values, 1))
    ReDim Observations(LBound(Values, 1) To UBound(Values, 1))
    For r = LBound(Values, 1) To UBound(Values, 1)
        SeriesName = Split(Series(r, 1), ".")
        If UBound(SeriesName) > 2 Then
            Observations(r) = SeriesName(UBound(SeriesName) - 1) & "." & SeriesName(UBound(SeriesName))
            ReDim Preserve SeriesName(UBound(SeriesName) - 2)
            SeriesNames(r) = VBA.Join(SeriesName, ".")
        End If
    Next r
    For c = LBound(Values, 2) To UBound(Values, 2)
        For r = LBound(Values, 1) To UBound(Values, 1)
            If Not Trim(CStr(Values(r, c))) = "" Then
                ts.WriteLine SeriesNames(r) & "," & Format(Dates(1, c), "yyyy-mm-dd")
                ts.WriteLine Observations(r)
                ts.WriteLine Trim(CStr(Values(r, c)))
                ts.WriteLine
            End If
        Next r
    Next c
    ts.Close
End Sub
Sub Macro2()
'
' Macro2 Macro
'

'
    Range("E26").Select
    ActiveCell.SpecialCells(xlLastCell).Select
    Range("F2").Select
End Sub
