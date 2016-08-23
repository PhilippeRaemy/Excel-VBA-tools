Attribute VB_Name = "BendingHelpers"
Option Explicit

Public Function GetWord(s As String, w As Integer) As String
Dim A As Variant
    A = Split(s, "/")
    Dim specChars As String: specChars = "-,;()"
    GetWord = A(w - 1)
    
    Dim i As Integer
    For i = 1 To Len(specChars)
        GetWord = Replace(GetWord, Mid(specChars, i, 1), " ")
    Next i
    While InStr(GetWord, "  ") > 0
        GetWord = Replace(GetWord, "  ", " ")
    Wend
    GetWord = Replace(Trim(GetWord), " ", "_")
End Function


