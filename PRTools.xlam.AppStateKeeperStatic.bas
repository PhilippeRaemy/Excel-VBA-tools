Attribute VB_Name = "AppStateKeeperStatic"
Option Explicit

Public Function NewAppStateKeeper() As AppStateKeeper
    Set NewAppStateKeeper = New AppStateKeeper
End Function

Public Function SetStatusBar(ByVal StatusMessage As String, ByVal seconds As Integer)
    Application.StatusBar = StatusMessage
    Application.OnTime DateAdd("s", seconds, Now()), "AppStateKeeperStatic.ResetStatusBar"
End Function

Public Function ResetStatusBar()
    Application.StatusBar = Null
End Function


