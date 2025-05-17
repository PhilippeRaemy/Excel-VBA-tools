Attribute VB_Name = "Styles"
Sub RemoveExtraneousStyles()
    Dim sty As Style, i As Long
    Dim re As New RegExp

    re.Pattern = ".*"
    For Each sty In ActiveWorkbook.Styles
        If Not sty.BuiltIn _
        And re.test(sty.Name) Then
            i = i + 1
            Debug.Print i, sty.Name,
            On Error Resume Next
            sty.Delete
            Debug.Print IIf(Err.Number = 0, "Deleted", "delete failed")
            On Error GoTo 0
        Else
            Debug.Print "keep " + ActiveWorkbook.Name, sty.Name, "BuiltIn:", sty.BuiltIn
        End If
    Next sty

End Sub

