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


