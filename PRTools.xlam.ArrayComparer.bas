Attribute VB_Name = "ArrayComparer"
Option Explicit

Public Function ArrayCompare(a As Variant, b As Variant) As Coordinate()
Dim r As Integer, c As Integer, cnt As Integer
Dim rc() As Coordinate
Dim coo As Coordinate: Set coo = New Coordinate

    If Not (LBound(a, 1) = LBound(b, 1) _
        And UBound(a, 1) = UBound(b, 1) _
        And LBound(a, 2) = LBound(b, 2) _
        And UBound(a, 2) = UBound(b, 2) _
    ) Then
        Err.Raise vbObjectError, "ArrayComparer.ArrayCompare", "Arrays are not of same dimensions"
    End If
    ReDim rc((1 + UBound(a, 1) - LBound(a, 1)) * (1 + UBound(a, 2) - LBound(a, 2)) - 1)
    For r = LBound(a, 1) To UBound(a, 1)
        For c = LBound(a, 2) To UBound(a, 2)
            If a(r, c) <> b(r, c) Then
                Set rc(cnt) = coo.NewCoordinate(r, c)
                cnt = cnt + 1
            End If
            On Error GoTo 0
        Next c
    Next r
    If cnt > 0 Then
        ReDim Preserve rc(cnt - 1)
        ArrayCompare = rc
    End If
End Function

Public Sub test()
Dim a(1 To 2, 1 To 3) As Variant
Dim b(1 To 2, 1 To 3) As Variant
Dim r As Integer, c As Integer
Dim i As Integer
Dim compareResults As Variant
Dim co As Coordinate, v As Variant
    
    For r = 1 To 2: For c = 1 To 3: a(r, c) = r * c: b(r, c) = r * c: Next: Next
    b(2, 2) = "foo"
    b(2, 3) = "bar"
    
    compareResults = ArrayCompare(a, b)
    For i = LBound(compareResults) To UBound(compareResults)
        Set co = compareResults(i)
        Debug.Print i; ":("; co.Row; ","; co.Column; ")"
    Next i
End Sub

