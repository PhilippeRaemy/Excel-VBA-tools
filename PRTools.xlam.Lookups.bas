Attribute VB_Name = "Lookups"
' ####################
' \\GVA0MS01\RAEMYP\Excel\Copy of PRTools.xlsm.Lookups.bas
' ####################
Option Explicit
Function LookupFirstNonNullValue(r As Range, Optional IgnoreWhiteSpace As Boolean = True) As Variant
Dim c As Range
  For Each c In r
    LookupFirstNonNullValue = c.value
    If Not IsEmpty(LookupFirstNonNullValue) Then
      If Trim(LookupFirstNonNullValue) <> "" Or Not IgnoreWhiteSpace Then
        Exit Function
      End If
    End If
  Next c
End Function
Function ConcatUntil(source As Range, stopper As Range, Optional Separator As String = " ") As String
Dim r As Integer, c As Integer

  For r = 1 To source.Columns.Count
    For c = 1 To source.Rows.Count
      If stopper.Cells(r, c).value <> 0 Then Exit Function
      ConcatUntil = ConcatUntil & IIf(ConcatUntil = "", "", Separator) & source.Cells(r, c).value
    Next c
  Next r
    
End Function
