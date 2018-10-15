Attribute VB_Name = "ErrorUtils"
Option Explicit
 
Public Enum ErrorNum
    InvalidProcedureCall = 5
    SubscriptOutOfRange = 9
    AlreadyInCollection = 457
End Enum

Public Sub ReRaiseError(ByVal eo As ErrObject, ByVal MethodName As String)
    Err.Raise eo.Number, eo.Source & vbNewLine & "    " & MethodName, _
            eo.Description, eo.HelpFile, eo.HelpContext
End Sub

Public Function MsgBoxAbortRetryIgnore(ByVal e As ErrObject, ByVal Title As String) As VbMsgBoxResult
    MsgBoxAbortRetryIgnore = _
    MsgBox("Error #" & e.Number & ": " & e.Description & vbNewLine & _
        vbNewLine & _
        "What is your wish and command?", _
        vbQuestion Or vbAbortRetryIgnore, Title, e.HelpFile, e.HelpContext)
End Function

Public Function MsgBoxRetryCancel(ByVal e As ErrObject, ByVal Title As String) As VbMsgBoxResult
    MsgBoxRetryCancel = _
    MsgBox("Error #" & e.Number & ": " & e.Description & vbNewLine & _
        vbNewLine & _
        "What is your wish and command?", _
        vbQuestion Or vbRetryCancel, Title, e.HelpFile, e.HelpContext)
End Function

Public Sub DisplayError(ByVal e As ErrObject, ByVal Title As String, _
    Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly _
)
    MsgBox "Error #" & e.Number & ": " & e.Description & vbNewLine & _
        vbNewLine & "From:" & vbNewLine & Err.Source & vbNewLine & vbNewLine & _
        "Aborting operation", _
        Buttons Or vbExclamation, Title, e.HelpFile, e.HelpContext
End Sub

Public Sub CheckIsInitialized(ByVal IsInitialized As Boolean, ByVal MethodName As String)
    If Not IsInitialized Then Err.Raise 17, MethodName, "Object not initialized"
End Sub

Public Sub RaiseError5AbstractClass(ByVal MethodName As String)
    Err.Raise ErrorNum.InvalidProcedureCall, MethodName, _
        "Invalid procedure call or argument - Cannot instantiate abstract class."
End Sub

Public Sub ClearCollection(ByVal col As Collection)
    Do While col.Count > 0: col.Remove 1:  Loop
End Sub
