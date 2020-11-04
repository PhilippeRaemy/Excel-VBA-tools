VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Coordinate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Row As Integer
Public column As Integer

Public Function NewCoordinate(ByVal Row As Integer, ByVal column As Integer) As Coordinate
    Set NewCoordinate = New Coordinate
    NewCoordinate.Row = Row
    NewCoordinate.column = column
End Function

