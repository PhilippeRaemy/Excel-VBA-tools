Attribute VB_Name = "MacroScheduler"
Option Explicit
Private Type Sched
    Procedure As String
    ScheduleTime As Date
    Active As Boolean
End Type
Private Schedules() As Sched
Private Inited As Boolean

Private Sub Init()
    Inited = True
    ReDim Schedules(10)
End Sub
Private Sub CleanUp()
    If Not Inited Then Init
    Dim i As Integer
    For i = LBound(Schedules) To UBound(Schedules)
        If Schedules(i).ScheduleTime < Now - 5 / 86400 Then Schedules(i).Active = False
    Next i
End Sub
Private Function AddSchedule(EarliestTime As Date, Procedure As String) As Integer
    Dim i As Integer, s As Sched
    Dim foundASpot As Boolean
    While Not foundASpot
        For i = LBound(Schedules) To UBound(Schedules): s = Schedules(i)
            If Not s.Active Then
                AddSchedule = i
                foundASpot = True
                Exit For
            End If
        Next i
        If Not foundASpot Then
            ReDim Preserve Schedules(UBound(Schedules) + 10)
        End If
    Wend
    Schedules(i).ScheduleTime = EarliestTime
    Schedules(i).Procedure = Procedure
    Schedules(i).Active = True
    AddSchedule = i
End Function

Public Function Schedule(EarliestTime As Date, Procedure As String, Optional LatestTime As Variant) As Boolean
    Dim ScheduleNumber As Integer
    If Not Inited Then Init
    ' CleanUp
    ScheduleNumber = AddSchedule(EarliestTime, Procedure)
    Application.OnTime EarliestTime, Procedure, LatestTime
    Debug.Print "Scheduled " & Procedure & " for " & VBA.Format(EarliestTime, "hh:mm:ss") & " as schedule #" & ScheduleNumber
    Schedule = True
End Function

Public Sub Cancel(Procedure As String)
    CancelImpl Procedure
End Sub

Public Sub CancelAll()
    CancelImpl ""
End Sub

Private Sub CancelImpl(Procedure As String)
    Dim i As Integer, s As Sched
    If Not Inited Then Init
    For i = LBound(Schedules) To UBound(Schedules): s = Schedules(i)
        If s.Active _
        And (Procedure = "" Or s.Procedure = Procedure) _
        Then
            On Error Resume Next
            Application.OnTime s.ScheduleTime, s.Procedure, Schedule:=False
            If Err.Number = 0 Then
                Debug.Print "Cancelled " & s.Procedure & " scheduled at " & VBA.Format(s.ScheduleTime, "hh:mm:ss")
            End If
            On Error GoTo 0
            Schedules(i).Active = False
        End If
    Next i
End Sub

Public Function ListSchedules() As String
Dim i As Integer, s As Sched
    If Not Inited Then Init
    For i = LBound(Schedules) To UBound(Schedules): s = Schedules(i)
        ListSchedules = ListSchedules & i & ") `" & s.Procedure & "` (" & IIf(s.Active, "", "in") & "active) scheduled at " & VBA.Format(s.ScheduleTime, "hh:mm:ss") & vbCrLf
    Next i
End Function

