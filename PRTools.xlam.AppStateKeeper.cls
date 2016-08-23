VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppStateKeeper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private PrevCalculation As XlCalculation
Private HasPrevCalculation As Boolean

Private PrevScreenUpdating As Boolean
Private HasPrevScreenUpdating As Boolean

Private PrevCursor As XlMousePointer
Private HasPrevCursor As Boolean

Public Function SetCalculation(ByVal Calculation As XlCalculation) As AppStateKeeper
    PrevCalculation = Application.Calculation
    Application.Calculation = Calculation
    HasPrevCalculation = True
    Set SetCalculation = Me
End Function
Public Function SetScreenUpdating(ByVal ScreenUpdating As Boolean) As AppStateKeeper
    PrevScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = ScreenUpdating
    HasPrevScreenUpdating = True
    Set SetScreenUpdating = Me
End Function
Public Function SetApplicationCursor(ByVal Cursor As XlMousePointer) As AppStateKeeper
    PrevCursor = Application.Cursor
    Application.Cursor = Cursor
    HasPrevCursor = True
    Set SetApplicationCursor = Me
End Function


Private Sub Class_Terminate()
    If HasPrevCalculation Then Application.Calculation = PrevCalculation
    If HasPrevScreenUpdating Then Application.ScreenUpdating = PrevScreenUpdating
    If HasPrevCursor Then Application.Cursor = PrevCursor
End Sub
