Attribute VB_Name = "ErrorUtilities"
Option Explicit

Public Sub ReraiseError(ByVal Err As ErrObject, _
    ByVal MethodName As String, _
    Optional ByVal Details As String _
)
    If Not IsMissing(Details) Then MethodName = MethodName & "(" & Details & ")"
    Err.Raise Err.Number, Err.Source & vbNewLine & _
        MethodName, _
        Err.Description
End Sub

Public Sub DisplayError(ByVal MyError As ErrObject, _
    ByVal MethodName As String, _
    Optional ByVal Details As String _
)
    Const Indent As String = vbNewLine & "    "
     
    If Not IsMissing(Details) Then MethodName = MethodName & "(" & Details & ")"
    MsgBox "Error #" & Err.Number & ": " & Err.Description & vbNewLine & _
            "From:" & vbNewLine & _
            Replace(Err.Source & vbNewLine & MethodName, vbNewLine, Indent), _
            vbOKOnly Or vbCritical, MethodName
End Sub

''' <summary>Displays a pop-up alert to ease debugging (if enabled).</summary>
Public Sub DisplayAlert(ByVal ModuleName As String)
    MsgBox "Pause to Ctrl-Break", vbOKOnly Or vbInformation, ModuleName
End Sub
