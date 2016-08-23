Attribute VB_Name = "ArrayHelpers"
Function ArrayDim(ByVal v As Variant) As Integer
Dim D As Integer
    If Not IsArray(v) Then Exit Function
    On Error GoTo ExitFct:
    
    While True
        D = UBound(v, ArrayDim + 1)
        ArrayDim = ArrayDim + 1
    Wend
    
ExitFct:

End Function

Function Concat(ByVal V1 As Variant, ByVal V2 As Variant) As Variant
Dim Ad1 As Integer, Ad2 As Integer
Dim A As Variant, i As Integer, K As Integer, Ix As Integer
Ad1 = ArrayDim(V1)
Ad2 = ArrayDim(V2)
If Ad1 <> Ad2 Then
    Err.Raise vbObjectError, "PRTools.Xlam.ArrayHelpers.Concat", "Arrays Must Have Same Number Dimensions"
End If
Select Case Ad1
    Case 0
        Concat = Array(V1, V2)
    Case 1
        A = Array()
        ReDim A(UBound(V1) - LBound(V1) + UBound(V2) - LBound(V2) + 1)
        For i = LBound(V1) To UBound(V1): A(Ix) = V1(i): Ix = Ix + 1: Next i
        For i = LBound(V2) To UBound(V2): A(Ix) = V2(i): Ix = Ix + 1: Next i
    Case 2
        If Not (LBound(V1, 2) = LBound(V2, 2) And UBound(V1, 2) = UBound(V2, 2)) Then
            Err.Raise vbObjectError, "PRTools.Xlam.ArrayHelpers.Concat", "Arrays' Second Dimension Does Not Match"
        End If
        A = Array()
        ReDim A(UBound(V1) - LBound(V1) + UBound(V2) - LBound(V2) + 1, LBound(V1, 2) To UBound(V1, 2))
        For i = LBound(V1, 1) To UBound(V1, 1): For j = LBound(V1, 2) To UBound(V1, 2): A(Ix, j) = V1(i, j): Ix = Ix + 1: Next j: Next i
        For i = LBound(V2, 1) To UBound(V2, 1): For j = LBound(V2, 2) To UBound(V2, 2): A(Ix, j) = V2(i, j): Ix = Ix + 1: Next j: Next i
    Case 3
        If Not ( _
                    LBound(V1, 2) = LBound(V2, 2) And UBound(V1, 2) = UBound(V2, 2) _
            And LBound(V1, 3) = LBound(V2, 3) And UBound(V1, 3) = UBound(V2, 3) _
        ) Then
            Err.Raise vbObjectError, "PRTools.Xlam.ArrayHelpers.Concat", "Arrays' Second Or Third Dimension Do Not Match"
        End If
        A = Array()
        ReDim A(UBound(V1) - LBound(V1) + UBound(V2) - LBound(V2) + 1, LBound(V1, 2) To UBound(V1, 2), LBound(V1, 3) To UBound(V1, 3))
        For i = LBound(V1, 1) To UBound(V1, 1): For j = LBound(V1, 2) To UBound(V1, 2): For K = LBound(V1, 3) To UBound(V1, 3): A(Ix, j, K) = V1(i, j, K): Ix = Ix + 1: Next K: Next j: Next i
        For i = LBound(V2, 1) To UBound(V2, 1): For j = LBound(V2, 2) To UBound(V2, 2): For K = LBound(V2, 3) To UBound(V2, 3): A(Ix, j, K) = V2(i, j, K): Ix = Ix + 1: Next K: Next j: Next i
    Case Else
        Err.Raise vbObjectError, "PRTools.Xlam.ArrayHelpers.Concat", "Arrays Of More Than 3 Dimensions Are Not Supported"
End Select

End Function

Public Function FlattenArray(ParamArray Parms() As Variant) As Variant
Dim A As Variant, i As Integer, p As Integer, Pp As Integer
Dim b As Variant, Pa As Variant
    Pa = Parms
    While LBound(Pa) = UBound(Pa)
        Pa = Pa(UBound(Pa))
    Wend
    A = Array()
    For p = LBound(Pa) To UBound(Pa)
        If IsArray(Pa(p)) Then
            b = FlattenArray(Pa(p))
            ReDim Preserve A(i + UBound(b) - UBound(b) + 1)
            For Pp = LBound(b) To UBound(b)
                A(i) = b(Pp)
                i = i + 1
            Next Pp
        Else
            ReDim Preserve A(i)
            A(i) = Pa(p)
            i = i + 1
        End If
    Next p
    FlattenArray = A
End Function

' Make Variant(Height, Width) From Variant(Height)(Width)
Public Function Make2DArray(ByVal i As Variant, Optional ByVal Width As Integer = -1) As Variant
On Error GoTo Err_Proc:
GoTo Proc
Err_Proc:
    Logger.error "Error {0} In {2}: {1}", Err.Number, Err.Description, "SimpleDataset.Make2DArray"
    Err.Raise Logger.ErrNumber, Logger.ErrSource, Logger.ErrDescription, Logger.ErrHelpFile, Logger.ErrHelpContext
    Resume
    Exit Function
Proc:
Dim r As Integer, c As Integer, A() As Variant
    If Not IsArray(i) Then Exit Function
    If UBound(i) = -1 Then Exit Function
    If Not IsArray(i(LBound(i))) Then
        Width = 0
    ElseIf Width = -1 Then
        Width = UBound(i(LBound(i)))
    End If
    ReDim A(LBound(i) To UBound(i), Width)
    For r = LBound(i) To UBound(i)
        For c = 0 To Width
            If IsArray(i(r)) Then
                If c <= UBound(i(r)) Then
                    A(r, c) = i(r)(c)
                End If
            Else
                A(r, c) = i(r)
            End If
        Next c
    Next r
    Make2DArray = A
End Function

Sub TestArrayDim()
    Dim s As Integer
    Dim V1(4) As Integer
    Dim V2(4, 4) As Integer
    Dim V3(4, 4, 4) As Integer
    Dim V4(4, 4, 4, 4) As Integer
    
    Debug.Print ArrayDim(s)
    Debug.Print ArrayDim(V1)
    Debug.Print ArrayDim(V2)
    Debug.Print ArrayDim(V3)
    Debug.Print ArrayDim(V4)
End Sub

Sub TestForEach()
    Dim A, E, i
    A = Array(1, 2, 3)
    Debug.Print "A=[";: For Each E In A: Debug.Print E; ",";: Next E: Debug.Print "]"
    For Each E In A: E = E * 2: Next E
    Debug.Print "A=[";: For Each E In A: Debug.Print E; ",";: Next E: Debug.Print "]"
    For i = LBound(A) To UBound(A): A(i) = A(i) * 2: Next i
    Debug.Print "A=[";: For Each E In A: Debug.Print E; ",";: Next E: Debug.Print "]"
End Sub

Function ArrayToString(A As Variant) As String
Dim i As Integer, j As Integer
Dim results() As String, r As Integer
    Select Case ArrayDim(A)
        Case 0: ArrayToString = "Array()"
        Case 1:
            For i = LBound(A) To UBound(A)
                ArrayToString = IIf(ArrayToString = "", "Array(", ArrayToString & ", ")
                If ArrayDim(A(i)) = 0 Then
                    ArrayToString = ArrayToString & CStr(A(i))
                Else
                    ArrayToString = ArrayToString & ArrayToString(A(i))
                End If
            Next i
            ArrayToString = ArrayToString & ")"
        Case 2:
            For i = LBound(A, 1) To UBound(A, 1)
                ArrayToString = IIf(ArrayToString = "", "Array(", ArrayToString & ", ")
                For j = LBound(A, 2) To UBound(A, 2)
                    ArrayToString = ArrayToString & IIf(j = LBound(A, 2), "( ", ", ")
                    If ArrayDim(A(i, i)) = 0 Then
                        ArrayToString = ArrayToString & CStr(A(i, j))
                    Else
                        ArrayToString = ArrayToString(A(i, i))
                    End If
                Next j
                ArrayToString = ArrayToString & ")"
            Next i
            ArrayToString = ArrayToString & ")"
        Case Else
            Err.Raise vbObjectError, "ArrayHelpers.ArrayToString", "3-dim arrays or higher not supported"
    End Select
End Function


Sub testArrayToString()
    Dim A(2, 3), i, j
    For i = LBound(A, 1) To UBound(A, 1)
        For j = LBound(A, 2) To UBound(A, 2)
            A(i, j) = i & "-" & j
        Next j
    Next i
    Debug.Print ArrayToString(A)
End Sub

Public Function ArrayContains(A As Variant, Value As Variant) As Boolean
    If Not IsArray(A) Then Exit Function
    Dim v As Variant
    For Each v In A
        If v = Value Then
            ArrayContains = True
            Exit Function
        End If
    Next v
End Function


