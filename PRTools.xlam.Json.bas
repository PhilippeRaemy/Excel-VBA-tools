Attribute VB_Name = "Json"
' VBA JSON parser, Backus-Naur form JSON parser based on RegEx v1.7.21
' Copyright (C) 2015-2020 omegastripes
' omegastripes@yandex.ru
' https://github.com/omegastripes/VBA-JSON-parser
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <http://www.gnu.org/licenses/>.

Option Explicit

' Need to include a reference to "Microsoft Scripting Runtime".

Private sBuffer As String
Private oTokens As Dictionary
Private oRegEx As Object
Private bMatch As Boolean
Private oChunks As Dictionary
Private oHeader As Dictionary
Private aData() As Variant
Private i As Long
Private sDelim As String
Private sTabChar As String
Private sLfChar As String
Private sSpcChar As String
    
Function Parse(ByVal sSample As String, sState As String) As Variant
    
    ' Input:
    ' sSample - source JSON string
    ' Output:
    ' vJson - created object or array to be returned as result
    ' sState - string Object|Array|Error depending on result
    
    Dim vJSON As Variant
    
    sBuffer = sSample
    Set oTokens = New Dictionary
    Set oRegEx = CreateObject("VBScript.RegExp")
    With oRegEx ' Patterns based on specification http://www.json.org/
        .Global = True
        .MultiLine = True
        .IgnoreCase = True ' Unspecified True, False, Null accepted
        .Pattern = "(?:'[^']*'|""(?:\\""|[^""])*"")(?=\s*[,\:\]\}])" ' Double-quoted string, unspecified quoted string
        Tokenize "s"
        .Pattern = "[+-]?(?:\d+\.\d*|\.\d+|\d+)(?:e[+-]?\d+)?(?=\s*[,\]\}])" ' Number, E notation number
        Tokenize "d"
        .Pattern = "\b(?:true|false|null)(?=\s*[,\]\}])" ' Constants true, false, null
        Tokenize "c"
        .Pattern = "\b[A-Za-z_]\w*(?=\s*\:)" ' Unspecified non-double-quoted property name accepted
        Tokenize "n"
        .Pattern = "\s+"
        sBuffer = .Replace(sBuffer, "") ' Remove unnecessary spaces
        .MultiLine = False
        Do
            bMatch = False
            .Pattern = "<\d+(?:[sn])>\:<\d+[codas]>" ' Object property structure
            Tokenize "p"
            .Pattern = "\{(?:<\d+p>(?:,<\d+p>)*)?,?\}" ' Object structure
            Tokenize "o"
            .Pattern = "\[(?:<\d+[codas]>(?:,<\d+[codas]>)*)?,?\]" ' Array structure
            Tokenize "a"
        Loop While bMatch
        .Pattern = "^<\d+[oa]>$" ' Top level object structure, unspecified array accepted
        If .test(sBuffer) And oTokens.Exists(sBuffer) Then
            sDelim = Mid(1 / 2, 2, 1)
            Retrieve sBuffer, vJSON
            sState = IIf(IsObject(vJSON), "Object", "Array")
        Else
            vJSON = Null
            sState = "Error"
        End If
    End With
    Set oTokens = Nothing
    Set oRegEx = Nothing
    
    Set Parse = vJSON
End Function

Private Sub Tokenize(sType)
    
    Dim aContent() As String
    Dim lCopyIndex As Long
    Dim i As Long
    Dim sKey As String
    
    With oRegEx.Execute(sBuffer)
        If .Count = 0 Then Exit Sub
        ReDim aContent(0 To .Count - 1)
        lCopyIndex = 1
        For i = 0 To .Count - 1
            With .item(i)
                sKey = "<" & oTokens.Count & sType & ">"
                oTokens(sKey) = .Value
                aContent(i) = Mid(sBuffer, lCopyIndex, .FirstIndex - lCopyIndex + 1) & sKey
                lCopyIndex = .FirstIndex + .Length + 1
            End With
        Next
    End With
    sBuffer = VBA.Join(aContent, "") & Mid(sBuffer, lCopyIndex, Len(sBuffer) - lCopyIndex + 1)
    bMatch = True
    
End Sub

Private Sub Retrieve(sTokenKey, vTransfer)
    
    Dim sTokenValue As String
    Dim sName As Variant
    Dim vValue As Variant
    Dim aTokens() As String
    Dim i As Long
    
    sTokenValue = oTokens(sTokenKey)
    With oRegEx
        .Global = True
        Select Case Left(Right(sTokenKey, 2), 1)
            Case "o"
                Set vTransfer = New Dictionary
                aTokens = Split(sTokenValue, "<")
                For i = 1 To UBound(aTokens)
                    Retrieve "<" & Split(aTokens(i), ">", 2)(0) & ">", vTransfer
                Next
            Case "p"
                aTokens = Split(sTokenValue, "<", 4)
                Retrieve "<" & Split(aTokens(1), ">", 2)(0) & ">", sName
                Retrieve "<" & Split(aTokens(2), ">", 2)(0) & ">", vValue
                If IsObject(vValue) Then
                    Set vTransfer(sName) = vValue
                Else
                    vTransfer(sName) = vValue
                End If
            Case "a"
                aTokens = Split(sTokenValue, "<")
                If UBound(aTokens) = 0 Then
                    vTransfer = Array()
                Else
                    ReDim vTransfer(0 To UBound(aTokens) - 1)
                    For i = 1 To UBound(aTokens)
                        Retrieve "<" & Split(aTokens(i), ">", 2)(0) & ">", vValue
                        If IsObject(vValue) Then
                            Set vTransfer(i - 1) = vValue
                        Else
                            vTransfer(i - 1) = vValue
                        End If
                    Next
                End If
            Case "n"
                vTransfer = sTokenValue
            Case "s"
                vTransfer = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace( _
                    Mid(sTokenValue, 2, Len(sTokenValue) - 2), _
                    "\""", """"), _
                    "\\", "\" & vbNullChar), _
                    "\/", "/"), _
                    "\b", Chr(8)), _
                    "\f", Chr(12)), _
                    "\n", vbLf), _
                    "\r", vbCr), _
                    "\t", vbTab)
                .Global = False
                .Pattern = "\\u[0-9a-fA-F]{4}"
                Do While .test(vTransfer)
                    vTransfer = .Replace(vTransfer, ChrW(("&H" & Right(.Execute(vTransfer)(0).Value, 4)) * 1))
                Loop
                vTransfer = Replace(vTransfer, "\" & vbNullChar, "\")
            Case "d"
                vTransfer = CDbl(Replace(sTokenValue, ".", sDelim))
            Case "c"
                Select Case LCase(sTokenValue)
                    Case "true"
                        vTransfer = True
                    Case "false"
                        vTransfer = False
                    Case "null"
                        vTransfer = Null
                End Select
        End Select
    End With
    
End Sub

Function Serialize(vJSON As Variant, Optional sTab As String = vbTab) As String
    
    If sTab = "" Then
        sTabChar = ""
        sLfChar = ""
        sSpcChar = ""
    Else
        sTabChar = sTab
        sLfChar = vbCrLf
        sSpcChar = " "
    End If
    Set oChunks = New Dictionary
    SerializeElement vJSON, ""
    Serialize = VBA.Join(oChunks.Items(), "")
    Set oChunks = Nothing
    
End Function

Private Sub SerializeElement(vElement As Variant, ByVal sIndent As String)
    
    Dim aKeys() As Variant
    Dim i As Long
    
    With oChunks
        Select Case VarType(vElement)
            Case vbObject
                If Not TypeOf vElement Is Dictionary Then
                    .item(.Count) = "{}"
                ElseIf vElement.Count = 0 Then
                    .item(.Count) = "{}"
                Else
                    .item(.Count) = "{" & sLfChar
                    aKeys = vElement.Keys
                    For i = 0 To UBound(aKeys)
                        .item(.Count) = sIndent & sTabChar & """" & EscapeJsonString(aKeys(i)) & """" & ":" & sSpcChar
                        SerializeElement vElement(aKeys(i)), sIndent & sTabChar
                        If Not (i = UBound(aKeys)) Then .item(.Count) = ","
                        .item(.Count) = sLfChar
                    Next
                    .item(.Count) = sIndent & "}"
                End If
            Case Is >= vbArray
                If UBound(vElement) = -1 Then
                    .item(.Count) = "[]"
                Else
                    .item(.Count) = "[" & sLfChar
                    For i = 0 To UBound(vElement)
                        .item(.Count) = sIndent & sTabChar
                        SerializeElement vElement(i), sIndent & sTabChar
                        If Not (i = UBound(vElement)) Then .item(.Count) = "," 'sResult = sResult & ","
                        .item(.Count) = sLfChar
                    Next
                    .item(.Count) = sIndent & "]"
                End If
            Case vbInteger, vbLong
                .item(.Count) = vElement
            Case vbSingle, vbDouble
                .item(.Count) = Replace(vElement, ",", ".")
            Case vbNull, vbError
                .item(.Count) = "null"
            Case vbBoolean
                .item(.Count) = IIf(vElement, "true", "false")
            Case Else
                .item(.Count) = """" & EscapeJsonString(vElement) & """"
        End Select
    End With
    
End Sub

Private Function EscapeJsonString(s)
    
    EscapeJsonString = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(s, _
        "\", "\\"), _
        """", "\"""), _
        "/", "\/"), _
        Chr(8), "\b"), _
        Chr(12), "\f"), _
        vbLf, "\n"), _
        vbCr, "\r"), _
        vbTab, "\t")
    
End Function

Function ToYaml(vJSON As Variant) As String
    
    Select Case VarType(vJSON)
        Case vbObject, Is >= vbArray
            Set oChunks = New Dictionary
            ToYamlElement vJSON, ""
            oChunks.Remove 0
            ToYaml = VBA.Join(oChunks.Items(), "")
            Set oChunks = Nothing
        Case vbNull, vbError
            ToYaml = "Null"
        Case vbBoolean
            ToYaml = IIf(vJSON, "True", "False")
        Case Else
            ToYaml = CStr(vJSON)
    End Select
    
End Function

Private Sub ToYamlElement(vElement As Variant, ByVal sIndent As String)
    
    Dim aKeys() As Variant
    Dim i As Long
    
    With oChunks
        Select Case VarType(vElement)
            Case vbObject
                If Not TypeOf vElement Is Dictionary Then
                    .item(.Count) = "''"
                ElseIf vElement.Count = 0 Then
                    .item(.Count) = "''"
                Else
                    .item(.Count) = vbCrLf
                    aKeys = vElement.Keys
                    For i = 0 To UBound(aKeys)
                        .item(.Count) = sIndent & aKeys(i) & ": "
                        ToYamlElement vElement(aKeys(i)), sIndent & "    "
                        If Not (i = UBound(aKeys)) Then .item(.Count) = vbCrLf
                    Next
                End If
            Case Is >= vbArray
                If UBound(vElement) = -1 Then
                    .item(.Count) = "''"
                Else
                    .item(.Count) = vbCrLf
                    For i = 0 To UBound(vElement)
                        .item(.Count) = sIndent & i & ": "
                        ToYamlElement vElement(i), sIndent & "    "
                        If Not (i = UBound(vElement)) Then .item(.Count) = vbCrLf
                    Next
                End If
            Case vbNull, vbError
                .item(.Count) = "Null"
            Case vbBoolean
                .item(.Count) = IIf(vElement, "True", "False")
            Case Else
                .item(.Count) = CStr(vElement)
        End Select
    End With
    
End Sub

Sub ToArray(vJSON As Variant, aRows() As Variant, aHeader() As Variant)
    
    ' Input:
    ' vJSON - Array or Object which contains rows data
    ' Output:
    ' aRows - 2d array representing JSON data
    ' aHeader - 1d array of property names
    
    Dim sName As Variant
    
    Set oHeader = New Dictionary
    Select Case VarType(vJSON)
        Case vbObject
            If vJSON.Count > 0 Then
                ReDim aData(0 To vJSON.Count - 1, 0 To 0)
                oHeader("#") = 0
                i = 0
                For Each sName In vJSON.Keys
                    aData(i, 0) = sName
                    ToArrayElement vJSON(sName), ""
                    i = i + 1
                Next
            Else
                ReDim aData(0 To 0, 0 To 0)
            End If
        Case Is >= vbArray
            If UBound(vJSON) >= 0 Then
                ReDim aData(0 To UBound(vJSON), 0 To 0)
                For i = 0 To UBound(vJSON)
                    ToArrayElement vJSON(i), ""
                Next
            Else
                ReDim aData(0 To 0, 0 To 0)
            End If
        Case Else
            ReDim aData(0 To 0, 0 To 0)
            aData(0, 0) = vJSON
    End Select
    aHeader = oHeader.Keys()
    Set oHeader = Nothing
    aRows = aData
    Erase aData
    
End Sub

Private Sub ToArrayElement(vElement As Variant, sFieldName As String)
    
    Dim sName As Variant
    Dim j As Long
    
    Select Case VarType(vElement)
        Case vbObject ' Collection of objects
            For Each sName In vElement.Keys
                ToArrayElement vElement(sName), sFieldName & IIf(sFieldName = "", "", ".") & sName
            Next
        Case Is >= vbArray  ' Collection of arrays
            For j = 0 To UBound(vElement)
                ToArrayElement vElement(j), sFieldName & "[" & j & "]"
            Next
        Case Else
            If Not oHeader.Exists(sFieldName) Then
                oHeader(sFieldName) = oHeader.Count
                If UBound(aData, 2) < oHeader.Count - 1 Then ReDim Preserve aData(0 To UBound(aData, 1), 0 To oHeader.Count - 1)
            End If
            j = oHeader(sFieldName)
            aData(i, j) = vElement
    End Select
    
End Sub

Sub Flatten(vJSON As Variant, vResult As Variant)
    
    ' Input:
    ' vJSON - Array or Object which contains JSON data
    ' Output:
    ' oResult - Flatten JSON data object
    
    Set oChunks = New Dictionary
    FlattenElement vJSON, ""
    Set vResult = oChunks
    Set oChunks = Nothing
    
End Sub

Private Sub FlattenElement(vElement As Variant, sProperty As String)
    
    Dim vKey
    Dim i As Long
    
    Select Case True
        Case TypeOf vElement Is Dictionary
            If vElement.Count > 0 Then
                For Each vKey In vElement.Keys
                    FlattenElement vElement(vKey), IIf(sProperty <> "", sProperty & "." & vKey, vKey)
                Next
            End If
        Case IsObject(vElement)
        Case IsArray(vElement)
            For i = 0 To UBound(vElement)
                FlattenElement vElement(i), sProperty & "[" & i & "]"
            Next
        Case Else
            oChunks(sProperty) = vElement
    End Select
    
End Sub

Sub Unflatten(oFlatten, vJSON, bSuccess)
    
    ' Input:
    ' oFlatten - source dictionary containing JSON data
    ' Output:
    ' vJSON - created object or array to be returned as result
    ' bSuccess - boolean indicating successful completion
    
    Dim sPath
    Dim vValue
    Dim aQualifiers
    Dim lNextLevel
    
    bSuccess = TypeOf oFlatten Is Dictionary
    If Not bSuccess Then Exit Sub
    For Each sPath In oFlatten.Keys
        If IsObject(oFlatten(sPath)) Then
            Set vValue = oFlatten(sPath)
        Else
            vValue = oFlatten(sPath)
        End If
        If Left(sPath, 1) <> "[" And Left(sPath, 1) <> "." Then
            sPath = "." & sPath
        End If
        aQualifiers = Split(Replace(Replace(sPath, ".", vbNullChar), "[", vbNullChar), vbNullChar)
        lNextLevel = 1
        UnflattenElement vJSON, lNextLevel, aQualifiers, vValue, bSuccess
        If Not bSuccess Then Exit Sub
    Next
    
End Sub

Private Sub UnflattenElement(vParent, lNextLevel, aQualifiers, vValue, bSuccess)
    
    Dim vNextQualifier
    Dim sNum
    Dim vChild
    
    bSuccess = False
    If lNextLevel > UBound(aQualifiers) Then
        If IsObject(vValue) Then
            Set vParent = vValue
        Else
            vParent = vValue
        End If
        bSuccess = True
        Exit Sub
    End If
    vNextQualifier = aQualifiers(lNextLevel)
    If Right(vNextQualifier, 1) = "]" Then
        sNum = Left(vNextQualifier, Len(vNextQualifier) - 1)
        If IsNumeric(sNum) Then
            vNextQualifier = CLng(sNum)
        End If
    End If
    If VarType(vNextQualifier) = vbLong Then
        If VarType(vParent) = vbEmpty Then
            vParent = Array()
        ElseIf Not IsArray(vParent) Then
            Exit Sub
        End If
        If UBound(vParent) < vNextQualifier Then
            ReDim Preserve vParent(vNextQualifier)
        End If
    Else
        If VarType(vParent) = vbEmpty Then
            Set vParent = New Dictionary
        ElseIf Not IsObject(vParent) Then
            Exit Sub
        ElseIf Not TypeOf vParent Is Dictionary Then
            Exit Sub
        End If
    End If
    If IsObject(vParent(vNextQualifier)) Then
        Set vChild = vParent(vNextQualifier)
    Else
        vChild = vParent(vNextQualifier)
    End If
    UnflattenElement vChild, lNextLevel + 1, aQualifiers, vValue, bSuccess
    If Not bSuccess Then
        Exit Sub
    End If
    If IsObject(vChild) Then
        Set vParent(vNextQualifier) = vChild
    Else
        vParent(vNextQualifier) = vChild
    End If
    bSuccess = True
    
End Sub

