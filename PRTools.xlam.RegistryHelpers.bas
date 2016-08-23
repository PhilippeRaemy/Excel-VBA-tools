Attribute VB_Name = "RegistryHelpers"
Option Explicit

Public Enum RegistryDataType
    REG_SZ
    REG_DWORD
    REG_BINARY
    REG_EXPAND_SZ
End Enum

Private Function RegistryDataTypeToString(dt As RegistryDataType) As String
    Select Case dt
        Case REG_SZ: RegistryDataTypeToString = "REG_SZ"
        Case REG_DWORD: RegistryDataTypeToString = "REG_DWORD"
        Case REG_BINARY: RegistryDataTypeToString = "REG_BINARY"
        Case REG_EXPAND_SZ: RegistryDataTypeToString = "REG_EXPAND_SZ"
    End Select
End Function

Public Function ReadRegistryValue(ByVal path As String) As Variant
Dim WshShell As New WshShell
    On Error Resume Next
    
    ReadRegistryValue = WshShell.RegRead(path)
    On Error GoTo 0
    
End Function


Public Function TestIfKeyExists(ByVal path As String) As Boolean
Dim WshShell As New WshShell
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

Public Sub WriteRegistryValue(ByVal path As String, ByVal Value As Variant, Optional ByVal DataType As RegistryDataType)
Dim WshShell As New WshShell
Dim DataTypeName As String
    DataTypeName = RegistryDataTypeToString(DataType)
    If DataTypeName = "" Then
        WshShell.RegWrite path, Value
    Else
        WshShell.RegWrite path, Value, DataTypeName
    End If
End Sub

Public Sub DeleteRegistryValue(ByVal path As String)
Dim WshShell As New WshShell
    WshShell.RegDelete path
End Sub
