Attribute VB_Name = "HouseKeeping"
Option Explicit

Sub DeleteEntireBottomAndRight()
'
' DeleteEntireBottomAndRight Macro: delete entire rows below and entire columns right of selection, to reset the last used cell pointer
'
'

Dim selected_cell As Range, selected_address As String
Dim column As String, i As Integer
Dim name_chunks As Variant
Dim wb As Workbook
    
    Set selected_cell = Selection
    selected_address = selected_cell.Address
    
    For i = 1 To Len(selected_address)
        Select Case UCase(Mid(selected_address, i, 1))
            Case "A" To "Z": column = column & UCase(Mid(selected_address, i, 1))
        End Select
    Next i
    
    With AppStateKeeperStatic.NewAppStateKeeper.SetScreenUpdating(False)
        Debug.Print Selection.Worksheet.Name
        Debug.Print Selection.Worksheet.UsedRange.Address, selected_cell.Address, selected_cell.Row, column, "->",
        Rows(selected_cell.Row & ":1048576").Select
        Selection.Delete Shift:=xlUp
        
        columns(column & ":XFD").Select
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.Delete Shift:=xlToLeft
        Debug.Print Selection.Worksheet.UsedRange.Address
        
        
        ActiveWindow.Zoom = 100
        Range("A1").Select
    End With
    DoEvents
    
    If vbYes = MsgBox("Save?", vbYesNo Or vbQuestion, "Housekeeping") Then
        Set wb = Selection.Worksheet.parent
        name_chunks = Split(wb.FullName, ".")
        If Not IsNumeric(name_chunks(UBound(name_chunks) - 1)) Then
            ReDim Preserve name_chunks(UBound(name_chunks) + 1)
            name_chunks(UBound(name_chunks)) = name_chunks(UBound(name_chunks) - 1)
        End If
        name_chunks(UBound(name_chunks) - 1) = Format(Now, "hhmmss")
        wb.SaveAs VBA.Join(name_chunks, ".")
        Debug.Print "saved as " & VBA.Join(name_chunks, ".")
    End If
End Sub

