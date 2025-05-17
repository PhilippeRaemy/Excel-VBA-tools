Attribute VB_Name = "VB_ObjectModel"
Option Explicit
Const SevenZip = """C:\Program Files\7-Zip\7z.exe"""

Public Sub ExportCode()
    ExportCodeImpl
End Sub


Public Sub ExportVersionsFromFolder()
Dim File As Variant
Dim fso As New scripting.FileSystemObject
Const FolderName = "\\kstlon0fs01\Shared\KS&T Global Gas\Installation\XLA\"
Const TargetFolderName = "C:\dev\GlobalGasAnalytics\XLA\"
Const TargetVbaFolderName = "C:\dev\GlobalGasAnalytics\XLA\KochGlobalGas.vba\"
Dim FileObject As scripting.File
Dim TargetFileObject As scripting.File
Dim targetFolder As scripting.Folder
If Not fso.FolderExists(FolderName) Then fso.CreateFolder FolderName
Set targetFolder = fso.GetFolder(FolderName)
If Not fso.FolderExists(TargetVbaFolderName) Then fso.CreateFolder TargetVbaFolderName
Dim vbaTargetFolder As scripting.Folder: Set vbaTargetFolder = fso.GetFolder(TargetVbaFolderName)


Const RootFileName = "KochGlobalGas"
Dim wb As Workbook
Dim cmd As CmdBatch

    For Each File In Versions
        Debug.Print FolderName & File
        On Error Resume Next
        Do While True
            Err.Clear
            fso.CopyFile FolderName & File, TargetFolderName & "KochGlobalGas.xlam", True
            If Err.Number = 0 Then Exit Do
            DoEvents
        Loop
        On Error GoTo 0
        fso.DeleteFolder vbaTargetFolder.path, True
        fso.CreateFolder TargetVbaFolderName
        Set wb = Application.Workbooks.Open(TargetFolderName & "KochGlobalGas.xlam", ReadOnly:=True)
        ExportCodeImpl wb, TargetFolderName & "KochGlobalGas.vba\", RootFileName
        Set cmd = New CmdBatch
        cmd.AddCmd "git add ."
        cmd.AddCmd "git commit -m " & File
        cmd.Run TargetFolderName
    Next File
End Sub



Public Sub ExportCodeImpl(Optional ByVal pwb As Workbook = Nothing, Optional targetFolder As String, Optional RootFileName As String)
    Dim c As Integer, l As Integer
    Dim VBProj
    Dim Comp As Variant ' VbComponent
    Dim FileName As String
    Dim Extension As scripting.Dictionary
    Set Extension = New scripting.Dictionary
    Dim wb As Workbook
    Dim UnzippedFolder As String
    Extension.Add 1, ".bas"
    Extension.Add 2, ".cls"
    Extension.Add 3, ".frm"
    Extension.Add 100, ".ws.bas"
    
    If pwb Is Nothing Then
        Set wb = Application.ActiveWorkbook
    Else
        Set wb = pwb
    End If
    If Not wb Is Nothing Then
        Set VBProj = wb.VBProject
    Else
        Set VBProj = Application.VBE.ActiveVBProject
    End If
    Debug.Print VBProj.FileName
    For c = 1 To VBProj.VBComponents.Count
        Set Comp = VBProj.VBComponents(c)
        If RootFileName = "" Or targetFolder = "" Then
            FileName = VBProj.FileName & "." & Comp.Name & Extension(Comp.Type)
        Else
            FileName = targetFolder & RootFileName & "." & Comp.Name & Extension(Comp.Type)
        End If
        Comp.Export FileName
    Next c
        
    Dim cmd As CmdBatch: Set cmd = New CmdBatch
    If pwb Is Nothing And Not ActiveWorkbook Is Nothing Then
        UnzippedFolder = """" & ActiveWorkbook.FullName & ".unzipped"""
        cmd.AddCmd "rd /s /q " & UnzippedFolder
        cmd.AddCmd SevenZip & " x -r -y """ & ActiveWorkbook.FullName & """ * -o" & UnzippedFolder
    Else
        If Not pwb Is Nothing Then
            UnzippedFolder = """" & wb.FullName & ".unzipped"""
            cmd.AddCmd "rd /s /q " & UnzippedFolder
            cmd.AddCmd SevenZip & " x -r -y """ & wb.FullName & """ * -o" & UnzippedFolder
        End If
    End If
    If Not wb Is Nothing Then
        FileName = wb.VBProject.FileName
        wb.Close True
    End If
    If cmd.CmdLine <> "" Then
        If pwb Is Nothing And Not ActiveWorkbook Is Nothing Then
            cmd.AddRestartWorkbook ActiveWorkbook.FullName
            ActiveWorkbook.Close
        End If
        cmd.Run
        If pwb Is Nothing Then Application.Quit
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
Dim fso        As New scripting.FileSystemObject
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

    Set TStream = fso.OpenTextFile(FileName, ForWriting, True)
    
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

Function Versions() As Collection
    Set Versions = New Collection
    Versions.Add "KochGlobalGas_5_4.xlam"
    Versions.Add "KochGlobalGas_6_0_0.xlam"
    Versions.Add "KochGlobalGas_6_0.xlam"
    Versions.Add "KochGlobalGas_6_1.xlam"
    Versions.Add "KochGlobalGas_6_1_0.xlam"
    Versions.Add "KochGlobalGas_6_1_1.xlam"
    Versions.Add "KochGlobalGas_6_2_0.xlam"
    Versions.Add "KochGlobalGas_6_2.xlam"
    Versions.Add "KochGlobalGas_6_3.xlam"
    Versions.Add "KochGlobalGas_6_4.xlam"
    Versions.Add "KochGlobalGas_6_2Dev.xlam"
    Versions.Add "KochGlobalGas_6_4_1.xlam"
    Versions.Add "KochGlobalGas_6_5.xlam"
    Versions.Add "KochGlobalGas_6_5_1.xlam"
    Versions.Add "KochGlobalGas_6_6.xlam"
    Versions.Add "KochGlobalGas_6_7.xlam"
    Versions.Add "KochGlobalGas_6_8.xlam"
    Versions.Add "KochGlobalGas_7_0.xlam"
    Versions.Add "KochGlobalGas_7_1.xlam"
    Versions.Add "KochGlobalGas_7_2.xlam"
    Versions.Add "KochGlobalGas_7_3.xlam"
    Versions.Add "KochGlobalGas_7_4_win7.xlam"
    Versions.Add "KochGlobalGas_7_4.xlam"
    Versions.Add "KochGlobalGas_8_0_old.xlam"
    Versions.Add "KochGlobalGas_8_0.xlam"
    Versions.Add "KochGlobalGas_9_0.xlam"
    Versions.Add "KochGlobalGas_9_1.xlam"
    Versions.Add "KochGlobalGas_9_2.xlam"
    Versions.Add "KochGlobalGas_9_3_old.xlam"
    Versions.Add "KochGlobalGas_9_3.xlam"
    Versions.Add "KochGlobalGas_9_5.xlam"
    Versions.Add "KochGlobalGas_10_0.xlam"
    Versions.Add "KochGlobalGas_10_1.xlam"
    Versions.Add "KochGlobalGas_10_2.xlam"
    Versions.Add "KochGlobalGas_10_3.xlam"
    Versions.Add "KochGlobalGas_10_4.xlam"
    Versions.Add "KochGlobalGas_10_5.xlam"
    Versions.Add "KochGlobalGas_10_6.xlam"
    Versions.Add "KochGlobalGas_10_7.xlam"
    Versions.Add "KochGlobalGas_10_8.xlam"
    Versions.Add "KochGlobalGas_10_9.xlam"
    Versions.Add "KochGlobalGas_10_10.xlam"
    Versions.Add "KochGlobalGas_10_11.xlam"
    Versions.Add "KochGlobalGas_10_12.xlam"
    Versions.Add "KochGlobalGas_10_13.xlam"
    Versions.Add "KochGlobalGas_10_14.xlam"
    Versions.Add "KochGlobalGas_10_15.xlam"

End Function


