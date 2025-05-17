Attribute VB_Name = "Addins"
Option Explicit
Sub exploreAddins()
Dim ai As Addin
Debug.Print Now
Dim co As New scripting.Dictionary
Dim maxLen As Integer
Dim key As Variant

    For Each ai In Application.AddIns2
        If True Or LCase(ai.Name) Like "*kquerycloud*" Then
            co.Add ai.Name & " {0}| " & IIf(ai.Installed, "active   | ", "inactive | ") & ai.FullName, Len(ai.Name)
            If Len(ai.Name) > maxLen Then maxLen = Len(ai.Name)
        End If
    Next ai
    
    Debug.Print String(maxLen, "-") & " + -------- + " & String(60, "-")
    For Each key In co.Keys
        Debug.Print Replace(key, "{0}", String(maxLen - co(key), " "))
    Next key
    Debug.Print String(maxLen, "-") & " + -------- + " & String(60, "-")
End Sub
