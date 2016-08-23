Attribute VB_Name = "Styles"
Sub Macro1()
'
' Macro1 Macro
'
Dim sty As Style, i As Integer
For Each sty In ActiveWorkbook.Styles
  If Regex.Match(sty.Name, " \d+") Then
    i = i + 1
    Debug.Print i, sty.Name
    sty.Delete
  Else
    Debug.Print "keep ", sty.Name
  End If
Next sty

End Sub

