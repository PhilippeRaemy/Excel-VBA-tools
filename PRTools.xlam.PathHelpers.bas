Attribute VB_Name = "PathHelpers"
Option Explicit

Sub ExportToCSVAndSave()

Dim FullName As String: FullName = ThisWorkbook.FullName
Dim RootFileName As String: RootFileName = GetFileName(FullName)
Dim Folder As String: Folder = GetFolderName(FullName)
   
    Application.DisplayAlerts = False
    ThisWorkbook.SaveAs filename:=Folder & "\Files\" & RootFileName & "." & ".csv", FileFormat:=xlCSV, CreateBackup:=False
    ThisWorkbook.SaveAs filename:=Folder & "\Back\" & RootFileName & Format(Now, """.""yyyymmdd"".""") & GetFileExt(FullName), FileFormat:=xlCSV, CreateBackup:=False
    ThisWorkbook.SaveAs filename:=FullName, FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
    Application.DisplayAlerts = True
End Sub


Public Function GetFolderName(FileFullName As String) As String
Dim aFile As Variant
    aFile = Split(FileFullName, "\")
    ReDim Preserve aFile(UBound(aFile) - 1)
    
    GetFolderName = VBA.Join(aFile, "\")
End Function

Public Function GetFileNameExt(FileFullName As String) As String
Dim aFile As Variant
    aFile = Split(FileFullName, "\")
    GetFileNameExt = aFile(UBound(aFile))
End Function

Public Function GetFileExt(FileFullName As String) As String
Dim aFile As Variant
    aFile = Split(GetFileNameExt(FileFullName), ".")
    GetFileExt = aFile(UBound(aFile))
End Function

Public Function GetFileName(FileFullName As String) As String
Dim aFile As Variant
    aFile = Split(GetFileNameExt(FileFullName), ".")
    ReDim Preserve aFile(UBound(aFile) - 1)
    GetFileName = VBA.Join(aFile, ".")
End Function



Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    ActiveWorkbook.SaveAs filename:= _
        "P:\CurveBuilder\V1\Power\Copy of LTShaping.xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=True
End Sub
