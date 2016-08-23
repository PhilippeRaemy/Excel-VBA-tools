VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProgressStatusBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Text As String
Public Max As Integer
Public BarSize As Integer
Private mProgress As Integer
Private Const SMILIE = "(O_O)"
Private Const BLINKINGL = "(O_~)"
Private Const BLINKINGR = "(~_O)"
Private Const BLINKING = "(~_~)"
Private blinkingSince As Date

Private Sub Class_Initialize()
    Max = 100
    BarSize = 60
End Sub

Private Sub Class_Terminate()
    Application.StatusBar = False
End Sub

Private Function face() As String
    Static lastFace As String
    If DateDiff("s", blinkingSince, Now) >= 1 Then
        Select Case lastFace
            Case BLINKING, BLINKINGL, BLINKINGR
                lastFace = SMILIE
            Case Else
                Select Case Rnd()
                    Case Is < 0.025: lastFace = BLINKING
                    Case Is < 0.05: lastFace = BLINKINGL
                    Case Is < 0.075: lastFace = BLINKINGR
                    Case Else: lastFace = SMILIE
                End Select
        End Select
        blinkingSince = Now
    End If
    face = lastFace
End Function

Public Sub Progress(Optional Progress As Integer = -1)
Dim sb As String
    If Max <= 0 Then Exit Sub
    mProgress = IIf(Progress = -1, mProgress + 1, Progress)
    If Text <> "" Then
        sb = Text & Format(mProgress / Max * 100, "00.0") & "% done"
    End If
    If BarSize <> 0 Then
        Dim sprogress As Single: sprogress = mProgress / Max
        If sprogress <= 1 Then
            sb = sb & " - " & String(sprogress * BarSize, "_") + face() + String((1 - sprogress) * BarSize, "_")
        End If
    End If
    Application.StatusBar = sb
End Sub




