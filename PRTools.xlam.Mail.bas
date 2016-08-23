Attribute VB_Name = "Mail"
Option Explicit

' Private WithEvents myOlItems  As Outlook.Items

Public Sub ConnectInbox()
    Dim olApp As Outlook.Application
    Dim objNS As Outlook.Namespace
    Set olApp = Outlook.Application
    Set objNS = olApp.GetNamespace("MAPI")
    Dim myOlItems  As Outlook.Items
    Set myOlItems = objNS.GetDefaultFolder(olFolderInbox).Items
    Dim mi As MailItem
    Dim i As Long, D As Long, v As Long, startTime As Date
    startTime = Now
    For i = myOlItems.Count To 1 Step -1
        ' Debug.Print i, myOlItems(i).Subject
        If myOlItems(i).Subject = "General Warning/Error - CommodityXL FxAll Trade Transformer" Then
            If TypeName(myOlItems(i)) = "MailItem" Then
                Set mi = myOlItems(i)
                mi.Delete
                D = D + 1
                ' mi.Delete
            End If
        End If
        v = v + 1
        If v Mod 100 = 0 Then
            Debug.Print Now, "counter:"; i; " visited:"; v; " deleted:"; D; " vRate:"; Int(v / (Now - startTime) / 24); "/[h]"; " dRate:"; Int(D / (Now - startTime) / 24); "/[h]"
        End If
        DoEvents
    Next i
    
End Sub

Private Sub myOlItems_ItemAdd(ByVal item As Object)

On Error GoTo ErrorHandler

    Dim Msg As Outlook.MailItem

    If TypeName(item) = "MailItem" Then
        Set Msg = item

        MsgBox Msg.Subject
        MsgBox Msg.Body

    End If

ProgramExit:
    Exit Sub
ErrorHandler:
    MsgBox Err.Number & " - " & Err.Description
    Resume ProgramExit
End Sub


