Attribute VB_Name = "ButtonProcessing"
Option Explicit
Option Private Module
Private Const ModuleName    As String = "ButtonProcessing"

Public Const COMAddInName   As String = "PGSolutions.BetterRibbon"

Public Sub Button1_Processing(ByVal SourceName As String)
    On Error GoTo EH
    MsgBox "Activation message from Button1!", vbOKOnly Or vbInformation, SourceName
XT: Exit Sub
EH: ErrorUtils.ReRaiseError Err, ModuleName & ".Button1_Processing"
    Resume          ' for debugging only
End Sub

Public Sub Button2_Processing(ByVal SourceName As String)
    On Error GoTo EH
    MsgBox "Activation message from Button2!", vbOKOnly Or vbInformation, SourceName
XT: Exit Sub
EH: ErrorUtils.ReRaiseError Err, ModuleName & ".Button2_Processing"
    Resume          ' for debugging only
End Sub

Public Sub Button3_Processing(ByVal SourceName As String)
    On Error GoTo EH
    MsgBox "Activation message from Button3!", vbOKOnly Or vbInformation, SourceName
XT: Exit Sub
EH: ErrorUtils.ReRaiseError Err, ModuleName & ".Button3_Processing"
    Resume          ' for debugging only
End Sub

Public Function ToggleImage(ByVal IsPressed As Boolean) As String
    ToggleImage = IIf(IsPressed, "TagMarkComplete", "MarginsShowHide")
End Function

Public Function ShowImage(ByVal SelectedIndex As Integer) As Boolean
    ShowImage = ((SelectedIndex + 1) And 2) <> 0
End Function

Public Function ShowLabel(ByVal SelectedIndex As Integer) As Boolean
    ShowLabel = ((SelectedIndex + 1) And 1) <> 0
End Function

Public Sub SetImageAndLabel(ByVal SelectedIndex As Integer, ParamArray Arr() As Variant)
    On Error GoTo EH
    Dim v As Variant
    For Each v In Arr
        Dim button As RibbonButtonModel: Set button = v
        If Not button Is Nothing Then
            button.ShowImage = ShowImage(SelectedIndex)
            button.ShowLabel = ShowLabel(SelectedIndex)
            button.Invalidate
        End If
    Next v
XT: Exit Sub
EH: ErrorUtils.ReRaiseError Err, ModuleName & ".Button2_Processing"
    Resume          ' for debugging only
End Sub
