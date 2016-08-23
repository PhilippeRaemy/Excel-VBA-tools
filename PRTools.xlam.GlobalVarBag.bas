Attribute VB_Name = "GlobalVarBag"
Option Explicit
Private Bag As Scripting.Dictionary

Public Function Deposit(ByVal var As Variant) As String
    GlobalVarBag.Init
    Deposit = Mid(CreateObject("Scriptlet.TypeLib").GUID, 2, 36)
    If IsObject(var) Then
        Set Bag(Deposit) = var
    Else
        Let Bag(Deposit) = var
    End If
End Function

Public Function Consult(ByVal ref As String, ByVal default As Variant) As Variant
    GlobalVarBag.Init
    If Bag.Exists(ref) Then
        If IsObject(Bag(ref)) Then
            Set Consult = Bag(ref)
        Else
            Let Consult = Bag(ref)
        End If
    Else
        If IsObject(default) Then
            Set Consult = default
        Else
            Let Consult = default
        End If
    End If
End Function

Public Function Withdraw(ByVal ref As String, ByVal default As Variant) As Variant
    GlobalVarBag.Init
    If Bag.Exists(ref) Then
        If IsObject(Bag(ref)) Then
            Set Withdraw = Bag(ref)
        Else
            Let Withdraw = Bag(ref)
        End If
        Bag.Remove ref
    Else
        If IsObject(default) Then
            Set Withdraw = default
        Else
            Let Withdraw = default
        End If
    End If
End Function

Public Sub Clear()
    If Bag Is Nothing Then
        Set Bag = New Scripting.Dictionary
    Else
        Bag.RemoveAll
    End If
End Sub

Public Sub Init()
    If Bag Is Nothing Then Set Bag = New Scripting.Dictionary
End Sub
