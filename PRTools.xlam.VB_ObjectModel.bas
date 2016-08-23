Attribute VB_Name = "VB_ObjectModel"
Option Explicit

Public Sub ExportCode()
    CheckinCode Checkin:=False
End Sub
Public Sub Checkin()
    CheckinCode Checkin:=True
End Sub
Public Sub CheckinCode(Optional Checkin As Boolean)
    Dim c As Integer, l As Integer
    Dim wb As Workbook
    Dim VBProj
    Dim Extension As Scripting.Dictionary
    Set Extension = New Scripting.Dictionary
    Extension.Add 1, ".bas"
    Extension.Add 2, ".cls"
    Extension.Add 3, ".frm"
    Extension.Add 100, ".ws.bas"
    
    Dim ChangedFiles As Scripting.Dictionary
    Set ChangedFiles = New Scripting.Dictionary
    Dim FilesToCheckout As String
    Dim FilesToAdd As String
    
    Dim FSO As FileSystemObject: Set FSO = New FileSystemObject
    Dim FileName As String, filenameTfs As String
    Dim ts As Scripting.TextStream
    Dim code As String, oldcode As String
    Dim fileStatus As String
    
    Dim wshsh As WshShell: Set wshsh = New WshShell
    
    FileName = DocumentActiveWorkbook(wshsh, Checkin)
    If Not FileName = "" Then
        FilesToAdd = " """ & FileName & """"
    End If
    If Not FileName = "" Then
        FilesToAdd = " """ & FileName & """"
    End If
    
    Set wb = Application.ActiveWorkbook
    If Not wb Is Nothing Then
        Set VBProj = wb.VBProject
    Else
        Set VBProj = Application.VBE.ActiveVBProject
    End If
    Debug.Print VBProj.FileName
    Dim TempFileNameRoot As String: TempFileNameRoot = "f" & Format(Now, "yyyymmdd_hhmmss")
    Dim TempFileName As String: TempFileName = Environ("tmp") & "\" & TempFileNameRoot & ".tmp"
    For c = 1 To VBProj.VBComponents.Count
        Dim Comp As Variant ' VbComponent
        Set Comp = VBProj.VBComponents(c)
        FileName = VBProj.FileName & "." & Comp.Name & Extension(Comp.Type)
        filenameTfs = VBProj.FileName & "." & Comp.Name & ".*"
        If FSO.FileExists(TempFileName) Then FSO.DeleteFile (TempFileName)
        Comp.Export TempFileName
        Set ts = FSO.OpenTextFile(TempFileName)
        code = ts.ReadAll
        ts.Close
        fileStatus = "New"
        If FSO.FileExists(FileName) Then
            Set ts = FSO.OpenTextFile(FileName)
            oldcode = Replace(ts.ReadAll, Mid(FSO.GetFileName(FileName), 1, Len(FSO.GetFileName(FileName)) - Len(FSO.GetExtensionName(FileName)) - 1), TempFileNameRoot)
            ts.Close
            If oldcode = code Then
                fileStatus = "Same"
            Else
                fileStatus = "Changed"
                Debug.Print " file "; Comp.Name; " has changed";
                If (FSO.GetFile(FileName).Attributes And ReadOnly) = ReadOnly Then
                    ' possibly checked in TFS: try to checkout
                    FilesToCheckout = FilesToCheckout & " """ & filenameTfs & """"
                    Debug.Print " and will be checked-out";
                End If
                Debug.Print "."
            End If
        End If
        If Not fileStatus = "Same" Then
            ChangedFiles.Add Comp.Name, FileName
        End If
        If fileStatus = "New" Then
            Debug.Print " file "; Comp.Name; " is new."
            FilesToAdd = FilesToAdd & " """ & filenameTfs & """"
        End If
    Next c
    
    If Checkin And Not FilesToCheckout = "" Then
        Debug.Print FilesToCheckout
        wshsh.Run "tf.bat checkout" & FilesToCheckout, WshNormalFocus, True
    End If
    
    For c = 0 To ChangedFiles.Count - 1
        If FSO.FileExists(ChangedFiles.Items(c)) Then
            FSO.DeleteFile VBProj.FileName & "." & ChangedFiles.Keys(c) & ".*"
        End If
        VBProj.VBComponents(ChangedFiles.Keys(c)).Export ChangedFiles.Items(c)
    Next c
    
    If Checkin And Not FilesToAdd = "" Then
        wshsh.Run "tf.bat add" & FilesToAdd, WshNormalFocus, True
    End If
    
    If Checkin And Not wb Is Nothing Then
        FileName = wb.VBProject.FileName
        wb.Save
        wb.Close True
        Dim cmd As String
        cmd = "cmd.exe /c "
        cmd = cmd & "tf.bat checkin    """ & FSO.GetFile(FileName).ParentFolder.path & "\*"""
        cmd = cmd & " & tf.bat checkout """ & FileName & """"
        cmd = cmd & " & timeout 1"
        cmd = cmd & " & ""C:\Program Files\Microsoft Office 15\root\office15\EXCEL.EXE"" /x """ & FileName & """"
        wshsh.Run cmd, WshNormalFocus, False
        Application.Quit
    End If
    
End Sub


' ==============================================================
' * Please note that Microsoft provides programming examples
' * for illustration only, without warranty either expressed or implied,
' * including, but not limited to, the implied warranties of merchantability
' * and/or fitness for a particular purpose. Any use by you of the code provided
' * in this blog is at your own risk.
'===============================================================

Sub CheckIfVBAAccessIsOn()

'[HKEY_LOCAL_MACHINE/Software/Microsoft/Office/10.0/Excel/Security]
'"AccessVBOM"=dword:00000001
 
Dim strRegPath As String
strRegPath = "HKEY_CURRENT_USER\Software\Microsoft\Office\" & Application.Version & "\Excel\Security\AccessVBOM"

If TestIfKeyExists(strRegPath) = False Then
'     Dim WSHShell As Object
'     Set WSHShell = CreateObject("WScript.Shell")
'     WSHShell.RegWrite strRegPath, 3, "REG_DWORD"
     MsgBox "A change has been introduced into your registry configuration. Pease restart Excel."
     WriteVBS
     Application.Quit
End If

Dim VBAEditor As Object         'VBIDE.VBE
Dim VBProj        As Object         'VBIDE.VBProject
Dim tmpVBComp As Object         'VBIDE.VBComponent
Dim VBComp        As Object         'VBIDE.VBComponent
        
Set VBAEditor = Application.VBE
Set VBProj = Application.ActiveWorkbook.VBProject
     

Dim counter As Integer

For counter = 1 To VBProj.References.Count
    Debug.Print VBProj.References(counter).FullPath
    'Debug.Print VBProj.References(counter).Name
    Debug.Print VBProj.References(counter).Description
    Debug.Print "---------------------------------------------------"
Next
 
End Sub

Function TestIfKeyExists(ByVal path As String)
Dim WshShell As Object
Set WshShell = CreateObject("WScript.Shell")
On Error Resume Next
    WshShell.RegRead path
 
    If Err.Number <> 0 Then
         Err.Clear
         TestIfKeyExists = False
    Else
         TestIfKeyExists = True
    End If
    On Error GoTo 0
End Function

Sub WriteVBS()
Dim objFile         As Object
Dim objFSO            As Object
Dim codePath        As String

    codePath = Application.ActiveDocument.path & "\reg_setting.vbs"
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.OpenTextFile(codePath, 2, True)
    
    objFile.WriteLine (" On Error Resume Next")
    objFile.WriteLine ("")
    objFile.WriteLine ("Dim WshShell")
    objFile.WriteLine ("Set WshShell = CreateObject(""WScript.Shell"")")
    objFile.WriteLine ("")
    objFile.WriteLine ("MsgBox ""Click OK to complete the setup process.""")
    objFile.WriteLine ("")
    objFile.WriteLine ("Dim strRegPath")
    objFile.WriteLine ("Dim Application_Version")
    objFile.WriteLine ("Application_Version = """ & Application.Version & """")
    objFile.WriteLine ("strRegPath = ""HKEY_CURRENT_USER\Software\Microsoft\Office\"" & Application_Version & ""\Excel\Security\AccessVBOM""")
    objFile.WriteLine ("WScript.echo strRegPath")
    objFile.WriteLine ("WshShell.RegWrite strRegPath, 1, ""REG_DWORD""")
    objFile.WriteLine ("")
    objFile.WriteLine ("If Err.Code <> o Then")
    objFile.WriteLine ("     MsgBox ""Error"" & Chr(13) & Chr(10) & Err.Source & Chr(13) & Chr(10) & Err.Message")
    objFile.WriteLine ("End If")
    objFile.WriteLine ("")
    objFile.WriteLine ("WScript.Quit")
    
    objFile.Close
    Set objFile = Nothing
    Set objFSO = Nothing

'run the VBscript code
    Shell "cscript " & codePath, vbNormalFocus

End Sub
Public Function DocumentActiveWorkbook(wshsh As WshShell, Checkin As Boolean) As String
Dim wb As Workbook, ws As Worksheet, nm As Name, lo As listobject, cell As Range
Dim TStream    As TextStream
Dim FSO        As New Scripting.FileSystemObject
Dim FileName As String
Dim fCond    As FormatCondition
Dim vfCond     As Variant

' On Error Resume Next

    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Function
    
    FileName = wb.FullName & ".txt"
    If Checkin Then
        wshsh.Run "tf.bat checkout " & FileName, WshNormalFocus, True
    End If

    Set TStream = FSO.OpenTextFile(FileName, ForWriting, True)
    
    If wb Is Nothing Then Exit Function
    TStream.WriteLine strings.FormatString("Workbook :\t{0}", wb.Name)
    For Each nm In wb.Names
        If InStr(CStr(nm), "#") > 0 Then
            TStream.WriteLine strings.FormatString("Workbook Named Range :\t{0}\t{1}", nm.Name, CStr(nm))
        Else
            TStream.WriteLine strings.FormatString("Workbook Named Range :\t{0}\t{1}!{2}", nm.Name, nm.RefersToRange.Worksheet.Name, nm.RefersToRange.Address)
        End If
        On Error GoTo 0
    Next nm
    For Each ws In wb.Worksheets
        TStream.WriteLine strings.FormatString("Worksheet :\t{0}", ws.Name)
        For Each nm In ws.Names
            If InStr(CStr(nm), "#") > 0 Then
                TStream.WriteLine strings.FormatString("Worksheet Named Range :\t{0}\t{1}", nm.Name, CStr(nm))
            Else
                TStream.WriteLine strings.FormatString("Worksheet Named Range :\t{0}\t{1}!{2}", nm.Name, nm.RefersToRange.Worksheet.Name, nm.RefersToRange.Address)
            End If
        Next nm
        For Each lo In ws.ListObjects
            TStream.WriteLine strings.FormatString("Worksheet List object :\t{0}\t{1}!{2}", lo.Name, lo.Range.Worksheet.Name, lo.Range.Address)
        Next lo
        For Each cell In ws.UsedRange.Cells
            If cell.Formula <> "" Then
                TStream.WriteLine strings.FormatString("{0}!{1}\t{2}", ws.Name, cell.Address, cell.Formula)
            End If
            For Each vfCond In cell.FormatConditions
                TStream.WriteLine strings.FormatString("{0}!{1}\t {2}", ws.Name, cell.Address, FormatConditionToString(vfCond))
            Next vfCond
        Next cell
    Next ws
    TStream.Close
    DocumentActiveWorkbook = FileName

End Function

Function FormatConditionToString(ByVal fc As Object) As String
    On Error Resume Next
    FormatConditionToString = strings.FormatString("FormatCondition: {0} ", TypeName(fc))
    Dim t As String
    Select Case fc.Type
        Case Excel.XlFormatConditionType.xlAboveAverageCondition: t = "AboveAverageCondition"
        Case Excel.XlFormatConditionType.xlBlanksCondition:         t = "BlanksCondition"
        Case Excel.XlFormatConditionType.xlCellValue:             t = "CellValue"
        Case Excel.XlFormatConditionType.xlColorScale:            t = "ColorScale"
        Case Excel.XlFormatConditionType.xlDatabar:                 t = "Databar"
        Case Excel.XlFormatConditionType.xlErrorsCondition:         t = "ErrorsCondition"
        Case Excel.XlFormatConditionType.xlExpression:            t = "Expression"
        Case Excel.XlFormatConditionType.xlIconSets:                t = "IconSets"
        Case Excel.XlFormatConditionType.xlNoBlanksCondition:     t = "NoBlanksCondition"
        Case Excel.XlFormatConditionType.xlNoErrorsCondition:     t = "NoErrorsCondition"
        Case Excel.XlFormatConditionType.xlTextString:            t = "TextString"
        Case Excel.XlFormatConditionType.xlTimePeriod:            t = "TimePeriod"
        Case Excel.XlFormatConditionType.xlTop10:                 t = "Top10"
        Case Excel.XlFormatConditionType.xlUniqueValues:            t = "UniqueValues"
        Case Else: t = "unknown condition type"
    End Select
    FormatConditionToString = strings.FormatString("{0} type={1}", FormatConditionToString, t)
    FormatConditionToString = strings.FormatString("{0}, applies to {1}", FormatConditionToString, fc.AppliesTo.Address)
End Function
