Attribute VB_Name = "SqlGen"
Option Explicit

Public Function GenInsert(ByVal TableName As String, ByVal columns As Range, ByVal Values As Range) As String
    GenInsert = GenInsertHead(TableName, columns)
    
End Function

Public Function GenInsertHead(ByVal TableName As String, ByVal columns As Range) As String
    Dim i As Integer, cell As Range
    For Each cell In columns.Cells
        If i = 0 Then
            i = 1
            GenInsertHead = "insert into " & TableName & "("
        Else
            GenInsertHead = GenInsertHead & ", "
        End If
        GenInsertHead = GenInsertHead & cell.Value
    Next cell
    GenInsertHead = GenInsertHead & ")"
End Function

Public Function GenValues(lineSeparator As String, Values As Range) As String
    Dim i As Integer, cell As Range
    For Each cell In Values.Cells
        If i = 0 Then
            i = 1
            GenValues = lineSeparator & "("
        Else
            GenValues = GenValues & ", "
        End If
        GenValues = GenValues & ToSqlLiteral(cell.Value)
    Next cell
    GenValues = GenValues & ")"
End Function

Public Function ToSqlLiteral(Value As Variant) As String
    If IsEmpty(Value) Then
        ToSqlLiteral = "null"
    ElseIf TypeName(Value) = "Date" Then
        ToSqlLiteral = Format(Value, IIf(Value = Int(Value), "yyyy-mm-dd", "yyyy-mm-dd hh:mm:ss"))
    ElseIf IsNumeric(Value) Then
        ToSqlLiteral = CStr(Value)
    Else
        ToSqlLiteral = "'" & Replace(CStr(Value), "'", "''") & "'"
    End If
End Function
