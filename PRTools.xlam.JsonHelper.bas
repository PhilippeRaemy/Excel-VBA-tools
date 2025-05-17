Attribute VB_Name = "JsonHelper"
Private ParsedJsons As New scripting.Dictionary

Function GetJsonProperty(Json As String, property As String)
    Dim parsed As Object
    Set parsed = ParseJson(Json)
    Debug.Print parsed

End Function

Function ParseJson(JsonString As String) As Object
    If Not ParsedJsons.Exists(JsonString) Then
        Dim state As String
        ParsedJsons.Add JsonString, Json.Parse(JsonString, state)
    End If

    Set ParseJson = ParsedJsons(JsonString)
    
End Function

