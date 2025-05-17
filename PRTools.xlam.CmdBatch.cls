VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CmdBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private batch As String

Public Sub AddCmd(cmd As String)
    If batch = "" Then
        batch = "cmd.exe /c """ & cmd
    Else
        batch = batch & " & " & cmd
    End If
End Sub
Public Sub AddRestartWorkbook(FileName As String)
    AddCmd "timeout 1"
    AddCmd """C:\Program Files\Microsoft Office\root\office16\EXCEL.EXE"" /x """ & ActiveWorkbook.FullName & """"
End Sub

Public Property Get CmdLine() As String
    If batch <> "" Then CmdLine = batch & """"
End Property

Public Sub Run(Optional CurrentDirectory As String)
    Dim wshsh As New WshShell
    If Not CurrentDirectory = "" Then
        wshsh.CurrentDirectory = CurrentDirectory
    End If
    Dim cmd As String: cmd = CmdLine()
    Debug.Print cmd
    wshsh.Run cmd, WshNormalFocus, False
End Sub

