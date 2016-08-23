Attribute VB_Name = "ShapeHelper"
Option Explicit

Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Rows("2:2").RowHeight = 37.5
End Sub

Sub RemoveShapes()
    While ActiveSheet.Shapes.Count > 0
        ActiveSheet.Shapes(1).Delete
    Wend
End Sub
Sub ResizeShapes()
Dim sh As Shape
    For Each sh In ActiveSheet.Shapes
        sh.Top = sh.TopLeftCell.Top
        sh.LockAspectRatio = True
        sh.height = sh.TopLeftCell.height
    Next sh
End Sub

