Attribute VB_Name = "DateHelpers"
Option Explicit

Public Function HoursPerGasDay(ByVal GasDay As Date, ByVal TimeZone As String) As Integer
' reference : http://en.wikipedia.org/wiki/Daylight_saving_time_by_country

    HoursPerGasDay = 24

    Select Case TimeZone
        Case "CET", "BPT", "BST"
            If DatePart("w", GasDay, vbMonday) = 6 Then
                If DatePart("m", GasDay) <> DatePart("m", GasDay + 7) Then
                    Select Case DatePart("m", GasDay)
                        Case 10: HoursPerGasDay = 25
                        Case 3: HoursPerGasDay = 23
                    End Select
                End If
            End If
    End Select
        
End Function
