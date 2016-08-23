Attribute VB_Name = "MessageBoxer"
Option Explicit

Private MessageDisplayed As Boolean
Private Type MessageType
    Prompt As String
    Buttons As VbMsgBoxStyle
    Title As String
    HelpFile As String
    Context As String
End Type
Private Messages() As MessageType
Private MessageCount As Integer

Private Function NewMessage(ByVal Prompt As String, ByVal Buttons As VbMsgBoxStyle, ByVal Title As String, ByVal HelpFile As String, ByVal Context As String) As MessageType
    NewMessage.Prompt = Prompt
    NewMessage.Buttons = Buttons
    NewMessage.Title = Title
    NewMessage.HelpFile = HelpFile
    NewMessage.Context = Context
End Function


Public Function MessageBox(ByVal Prompt As String, ByVal Buttons As VbMsgBoxStyle, ByVal Title As String, ByVal HelpFile As String, ByVal Context As String) As VbMsgBoxStyle
    If (Buttons And vbAbortRetryIgnore = vbAbortRetryIgnore) _
    Or (Buttons And vbOKCancel = vbOKCancel) _
    Or (Buttons And vbRetryCancel = vbRetryCancel) _
    Or (Buttons And vbYesNo = vbYesNo) _
    Or (Buttons And vbYes = vbYes) _
    Or (Buttons And vbNo = vbNo) _
    Or (Buttons And vbYesNoCancel = vbYesNoCancel) _
    Then
       VBA.MsgBox MessageBoxer.MessageBox(Prompt, Buttons, Title, HelpFile, Context)
       Exit Function
    End If
    
    
    ReDim Preserve Messages(MessageCount): MessageCount = MessageCount + 1
    Messages(MessageCount) = NewMessage(Prompt, Buttons, Title, HelpFile, Context)
    If Not MessageDisplayed Then DisplayMessage
    MessageBox = vbOK
End Function


Private Sub DisplayMessage()
    MessageDisplayed = True
    Dim uMsg As MessageType, mMsg As MessageType
    Dim m As Integer
    uMsg = Messages(0)
    For m = 1 To UBound(Messages)
        mMsg = Messages(m)
        If mMsg.Buttons > uMsg.Buttons Then uMsg.Buttons = mMsg.Buttons
        uMsg.Title = IIf(uMsg.Title = "", mMsg.Title, "Multiple messages")
        uMsg.Prompt = uMsg.Prompt & vbCrLf & mMsg.Title & " : " & mMsg.Prompt
    Next m
    VBA.MsgBox uMsg.Prompt, uMsg.Buttons, uMsg.Title, uMsg.HelpFile, uMsg.Context
    MessageCount = 0
    MessageDisplayed = False
End Sub


