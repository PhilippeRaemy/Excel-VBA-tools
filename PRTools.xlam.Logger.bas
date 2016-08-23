Attribute VB_Name = "Logger"
' ####################
' \\GVA0MS01\RAEMYP\Excel\Copy of PRTools.xlsm.Logger.bas
' ####################
Option Explicit

Private Const Logging As Boolean = True
Public WorkBookName As String
Private FSO As FileSystemObject
Private pLogFileName As String
Private Inited As Boolean
Private lErrHelpContext As String
Private lErrHelpFile As String
Private lErrLastDllError As Long
Private lErrNumber As Long
Private lErrsource As String
Private lErrDescription  As String

Public Property Get ErrHelpContext() As String
    ErrHelpContext = lErrHelpContext
End Property
Public Property Get ErrHelpFile() As String
    ErrHelpFile = lErrHelpFile
End Property
Public Property Get ErrLastDllError() As Long
    ErrLastDllError = lErrLastDllError
End Property
Public Property Get ErrNumber() As Long
    ErrNumber = lErrNumber
End Property
Public Property Get ErrSource() As String
    ErrSource = lErrsource
End Property
Public Property Get ErrDescription() As String
    ErrDescription = lErrDescription
End Property


Public Property Get LogFileName()
    Dim ai As AddIn
    If Not Inited Then
        If WorkBookName = "" Then WorkBookName = ActiveWorkbook.Name
        If pLogFileName = "" Then pLogFileName = Environ("temp") & "\" & WorkBookName & "_" & Format(Now, "yyyymmdd") & ".log"
        For Each ai In AddIns
            If ai.Name = "PRTools.xlam" And ai.Installed And ai.IsOpen Then
                Application.Run "ExportCode"
            End If
        Next ai
        Inited = True
    End If
    LogFileName = pLogFileName
End Property

Public Sub error(Msg As String, ParamArray Parms() As Variant)
  logmsg Msg, Parms
End Sub
Public Sub log(Msg As String, ParamArray Parms() As Variant)
  logmsg Msg, Parms
End Sub
Private Sub logmsg(Msg As String, ParamArray Parms() As Variant)
    lErrHelpContext = Err.HelpContext
    lErrHelpFile = Err.HelpFile
    lErrLastDllError = Err.LastDllError
    lErrNumber = Err.Number
    lErrsource = Err.source
    lErrDescription = Err.Description
    
    If Not Logger.Logging Then Exit Sub
    Dim p As Variant, FileName As String
    p = Parms
    If UBound(Parms) >= 0 Then
        If IsArray(Parms(0)) Then
            p = Parms(0)
        End If
    End If
    Dim i As Integer, ts As TextStream
    For i = 0 To UBound(p)
         Msg = Replace(Msg, "{" & i & "}", CStr(p(i)))
    Next i
    FileName = LogFileName
    Debug.Print Msg
    If FSO Is Nothing Then Set FSO = New FileSystemObject
    
    Set ts = FSO.OpenTextFile(FileName, ForAppending, True)
    ts.Write Format(Now, "yyyymmdd hhmmss")
    ts.Write " - "
    ts.WriteLine Msg
  ts.Close
End Sub

Public Sub ViewLogFile()
    Shell "e.bat """ & LogFileName & """"
End Sub
