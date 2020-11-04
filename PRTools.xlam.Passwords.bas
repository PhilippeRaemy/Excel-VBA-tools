Attribute VB_Name = "Passwords"
Option Explicit
Sub PasswordBreaker(MinLength As Integer, MaxLength As Integer)
    MsgBox "Password is " & PasswordBreakerImpl(MinLength, MaxLength, "")
End Sub

Function PasswordBreakerImpl(MinLength As Integer, MaxLength As Integer, PasswordRoot As String, Optional alphanum As Integer = 0)

' code is wrong for it's either fully alphanum or fully non-alphanum
' plus, it takes ~60-80 miliseconds per password test which makes it 270+ hours to crack a 3 chars password, ~375`535 years to crack 6 chars...
    Static p As Long
    Dim c As Integer
    Dim password As String

    If alphanum = 0 Then
        PasswordBreakerImpl = PasswordBreakerImpl(MinLength, MaxLength, "", 1)
        If PasswordBreakerImpl <> "" Then Exit Function
        PasswordBreakerImpl = PasswordBreakerImpl(MinLength, MaxLength, "", -1)
        Exit Function
    End If
    
    On Error Resume Next
    Dim isalphanumeric As Boolean
    For c = 32 To 126
        Select Case c
            Case 48 To 57, 65 To 90, 95, 97 To 122
                isalphanumeric = True
            Case Else
                isalphanumeric = False
        End Select
        If (alphanum = 1 And isalphanumeric) _
        Or (alphanum = -1 And Not isalphanumeric) _
        Then
            password = PasswordRoot & Chr(c)
            If p Mod 1000 = 0 Then
                Debug.Print Now, p, password
                DoEvents
            End If
            p = p + 1
            If Len(password) >= MinLength Then
                ActiveSheet.Unprotect password
                If ActiveSheet.ProtectContents = False Then
                    PasswordBreakerImpl = password
                    Exit Function
                End If
            End If
            If Len(password) < MaxLength Then
                PasswordBreakerImpl = PasswordBreakerImpl(MinLength, MaxLength, password, alphanum)
                If PasswordBreakerImpl <> "" Then Exit Function
            End If
        End If
    Next c
End Function

