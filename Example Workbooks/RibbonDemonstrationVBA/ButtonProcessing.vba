Attribute VB_Name = "ButtonProcessing"
Option Explicit
Option Private Module
Private Const ModuleName    As String = "ButtonProcessing"

Public Sub Button1_Processing(SourceName As String)
    On Error GoTo EH
    MsgBox "Activation message from Button1!", vbOKOnly Or vbInformation, SourceName
XT: Exit Sub
EH: ErrorUtils.ReRaiseError Err, ModuleName & ".Button1_Processing"
    Resume          ' for debugging only
End Sub

Public Sub Button2_Processing(SourceName As String)
    On Error GoTo EH
    MsgBox "Activation message from Button2!", vbOKOnly Or vbInformation, SourceName
XT: Exit Sub
EH: ErrorUtils.ReRaiseError Err, ModuleName & ".Button2_Processing"
    Resume          ' for debugging only
End Sub

Public Sub Button3_Processing(SourceName As String)
    On Error GoTo EH
    MsgBox "Activation message from Button3!", vbOKOnly Or vbInformation, SourceName
XT: Exit Sub
EH: ErrorUtils.ReRaiseError Err, ModuleName & ".Button3_Processing"
    Resume          ' for debugging only
End Sub
