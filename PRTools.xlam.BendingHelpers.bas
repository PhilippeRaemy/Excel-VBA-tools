Attribute VB_Name = "BendingHelpers"
Option Explicit

Public Function GetSpecialWord(s As String, w As Integer) As String
Dim a As Variant
    a = Split(s, "/")
    Dim specChars As String: specChars = "-,;()"
    GetSpecialWord = a(w - 1)
    
    Dim i As Integer
    For i = 1 To Len(specChars)
        GetSpecialWord = Replace(GetSpecialWord, Mid(specChars, i, 1), " ")
    Next i
    While InStr(GetSpecialWord, "  ") > 0
        GetSpecialWord = Replace(GetSpecialWord, "  ", " ")
    Wend
    GetSpecialWord = Replace(Trim(GetSpecialWord), " ", "_")
End Function

Public Function JoinString(Separator As String, ParamArray words() As Variant) As String
    JoinString = VBA.Join("_", words)
End Function

