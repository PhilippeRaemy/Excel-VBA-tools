Attribute VB_Name = "RegexHelpers"
Option Explicit

Function ExtractPattern(s As String, p As String) As String
Dim re As New RegExp
Dim ma As MatchCollection
    
    re.Pattern = p
    Set ma = re.Execute(s)
    If ma.Count > 0 Then
        If ma(0).SubMatches.Count > 0 Then
            ExtractPattern = ma(0).SubMatches(0)
        End If
    End If
    

End Function
