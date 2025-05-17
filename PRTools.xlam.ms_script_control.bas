Attribute VB_Name = "ms_script_control"
Option Explicit

Sub test()
Dim ScriptControl As Object
    Set ScriptControl = New ScriptControl ' CreateObject("ScriptControl")
    Set ScriptControl = New MSScriptControl.ScriptControl
    ScriptControl.Language = "VBScript"
    Debug.Print (ScriptControl.Eval("2+5"))

End Sub
