Attribute VB_Name = "strings"
Option Explicit


Function FormatString(Format As String, ParamArray args() As Variant)
    Dim a As Integer, arguments As Variant
    arguments = args
    If UBound(arguments) - LBound(arguments) >= 0 Then
        If IsArray(arguments(LBound(arguments))) Then
            arguments = arguments(LBound(arguments))
        End If
    End If
    FormatString = Format
    For a = LBound(arguments) To UBound(arguments)
        FormatString = Replace(FormatString, "{" & a & "}", arguments(a))
    Next a
    FormatString = Replace(Replace(FormatString, "\t", vbTab), "\n", vbCrLf)
End Function

Function isEmptyOrBlank(s As Variant) As Boolean
    If IsEmpty(s) Then
        isEmptyOrBlank = True
    ElseIf IsObject(s) Then
        If s Is Nothing Then isEmptyOrBlank = True
    ElseIf Trim(CStr(s)) = "" Then
        isEmptyOrBlank = True
    End If
End Function
Function isEmptyOrBlankOrZero(s As Variant) As Boolean
    If IsEmpty(s) Then
        isEmptyOrBlankOrZero = True
    ElseIf IsObject(s) Then
        If s Is Nothing Then isEmptyOrBlankOrZero = True
    ElseIf Trim(CStr(s)) = "" Then
        isEmptyOrBlankOrZero = True
    ElseIf Trim(CStr(s)) = "0" Then
        isEmptyOrBlankOrZero = True
    End If
End Function

Public Function GetFormula(r As Range) As String
    GetFormula = r.Formula
End Function

Public Function GetRangeName(r As Range) As String

On Error Resume Next
Proc:
    Dim n As Name
    For Each n In Application.Names
        'Debug.Print n.name,
        'Debug.Print n.RefersToRange.Worksheet.name,
        'Debug.Print n.RefersToRange.Address,
        If n.RefersToRange.Address = r.Address _
        And n.RefersToRange.Worksheet.Name = r.Worksheet.Name Then
            If Err.Number = 0 Then
                GetRangeName = n.Name
                Exit Function
            Else
                'Debug.Print Err.Description
                Err.Clear
            End If
        End If
        Debug.Print
    Next n
End Function

Public Function Min(a As Long, b As Long) As Long
    If a < b Then Min = a Else Min = b
End Function
Public Function Max(a As Long, b As Long) As Long
    If a > b Then Max = a Else Max = b
End Function

Public Function SplitStringH(s As String, delimiter As String) As Variant
    SplitStringH = VBA.Split(s, delimiter)
End Function

Public Function SplitStringV(s As String, delimiter As String) As Variant
    Dim results, pivoted As Variant, i As Integer
    results = VBA.Split(s, delimiter)
    pivoted = Array()
    ReDim pivoted(0 To UBound(results), 0 To 0)
    For i = LBound(results) To UBound(results)
        pivoted(i, 0) = results(i)
    Next i
    SplitStringV = pivoted
   
End Function

Public Function MReplace(OldText As String, Searches As Range, Replaces As Range, Optional Compare As VbCompareMethod = VbCompareMethod.vbTextCompare)
    MReplace = MSubstitute(OldText, Searches, Replaces, Compare)
End Function

Public Function MSubstitute(OldText As String, Searches As Range, Replaces As Range, Optional Compare As VbCompareMethod = VbCompareMethod.vbTextCompare)
    Dim sc As Integer, sr As Integer
    Dim rc As Integer, rr As Integer
    sc = 1: sr = 1: rc = 1: rr = 1
    MSubstitute = OldText
    Do
        MSubstitute = Replace(MSubstitute, Searches.Cells(sr, sc).Value, Replaces.Cells(rr, rc).Value, Compare)
        sc = sc + 1: If sc > Searches.columns.Count Then sc = 1: sr = sr + 1: If sr > Searches.Rows.Count Then Exit Function
        rc = rc + 1: If rc > Replaces.columns.Count Then rc = 1: rr = rr + 1: If rr > Replaces.Rows.Count Then Exit Function
    Loop
End Function

Public Function DeduplicateWords(text As String, Separator As String, Optional Compare As VbCompareMethod = VbCompareMethod.vbTextCompare) As String
    Dim w As Variant, d As Scripting.Dictionary
    Set d = New Scripting.Dictionary
    For Each w In Split(text, Separator, Compare:=Compare)
        If Not d.Exists(w) Then
            DeduplicateWords = IIf(DeduplicateWords = "", "", DeduplicateWords & Separator) & w
            d.Add w, 0
        End If
    Next w
End Function
Public Function GetWord(text As String, Separator As String, Position As Integer) As String
    Dim a As Variant
    a = VBA.Split(text, Separator)
    Select Case Position - 1
        Case LBound(a) To UBound(a)
            GetWord = a(Position - 1)
    End Select
End Function
Public Function FindFirstOf(text As String, rng As Range) As String
Dim c As Range
    For Each c In rng
        If InStr(0, text, CStr(c.Value), vbTextCompare) > 0 Then
            FindFirstOf = CStr(c.Value)
            Exit Function
        End If
    Next c
End Function

Public Function RemoveWord(text As String, Separator As String, Position As Integer) As String
    Dim a As Variant
    a = VBA.Split(text, Separator)
    If Position = 0 Or Position > UBound(a) + 1 Or Position < -1 - UBound(a) Then
        RemoveWord = text
        Exit Function
    End If
    Dim b As Variant
    b = Array()
    ReDim b(UBound(a) - 1)
    Dim i As Integer, j As Integer
    j = -1
    For i = 0 To UBound(a)
        If Not (i + 1 = Position Or (Position < 0 And UBound(a) + 1 - i = -Position)) Then
            j = j + 1
            b(j) = a(i)
        End If
    Next i
    RemoveWord = VBA.Join(b, Separator)
End Function

Public Sub TestRemoveWord()
Dim a As Variant

a = Array( _
    Array("a,b,c,d,e", 0, "a,b,c,d,e"), _
    Array("a,b,c,d,e", 1, "b,c,d,e"), _
    Array("a,b,c,d,e", 2, "a,c,d,e"), _
    Array("a,b,c,d,e", 3, "a,b,d,e"), _
    Array("a,b,c,d,e", 4, "a,b,c,e"), _
    Array("a,b,c,d,e", 5, "a,b,c,d"), _
    Array("a,b,c,d,e", 6, "a,b,c,d,e"), _
    Array("a,b,c,d,e", -5, "b,c,d,e"), _
    Array("a,b,c,d,e", -4, "a,c,d,e"), _
    Array("a,b,c,d,e", -3, "a,b,d,e"), _
    Array("a,b,c,d,e", -2, "a,b,c,e"), _
    Array("a,b,c,d,e", -1, "a,b,c,d"), _
    Array("a,b,c,d,e", -6, "a,b,c,d,e") _
)
    Dim t As Variant
    Dim results As String
    For Each t In a
        results = RemoveWord(CStr(t(0)), ",", CInt(t(1)))
        Debug.Print "RemoveWord(""" & t(0) & """, " & t(1) & ")=" & results, results = CStr(t(2))
        Debug.Assert results = CStr(t(2))
    Next t
End Sub

