Attribute VB_Name = "ButtonProcessing"
'''=======================================================================================
'''                Copyright (c) 2017-2019 Pieter Geerkens
'''
'''     Licensed under the MIT Licence at:
'''             https://github.com/pgeerkens/PGSolutions.BetterRibbon/blob/dev/LICENSE
'''=======================================================================================
Option Explicit
Option Private Module
Private Const ModuleName    As String = "ButtonProcessing"

Public Function AlternateToggle(ByVal Factory As IModelServer, ByVal Mode As Boolean, _
        ByVal Model As ToggleModel, ByVal ToggleID As String, ByVal CheckBoxID As String _
) As Boolean
    On Error GoTo EH
    AlternateToggle = Not Mode
    
    Model.Detach
    Model.Attach IIf(AlternateToggle, ToggleID, CheckBoxID)
    Model.SetImage ThisWorkbook.ModelServer.NewImageObjectMso(ToggleImage(Model.IsPressed))
    Model.Invalidate
    Application.StatusBar = "Ready ...'"
    
XT: Exit Function
EH: ErrorUtils.ReRaiseError Err, ModuleName & ".AlternateToggle"
    Resume          ' for debugging only
End Function

Public Sub SetImageAndLabel(ByVal SelectedIndex As Integer, ParamArray Arr() As Variant)
    On Error GoTo EH
    Dim v As Variant
    For Each v In Arr
        Dim button As ButtonModel: Set button = v
        If Not button Is Nothing Then
            button.ShowImage = ShowImage(SelectedIndex)
            button.ShowLabel = ShowLabel(SelectedIndex)
            button.Invalidate
        End If
    Next v
XT: Exit Sub
EH: ErrorUtils.ReRaiseError Err, ModuleName & ".SetImageAndLabel"
    Resume          ' for debugging only
End Sub

Public Function ToggleImage(ByVal IsPressed As Boolean) As String
    ToggleImage = IIf(IsPressed, "TagMarkComplete", "MarginsShowHide")
End Function

Private Function ShowImage(ByVal SelectedIndex As Integer) As Boolean
    ShowImage = ((SelectedIndex + 1) And 2) <> 0
End Function

Private Function ShowLabel(ByVal SelectedIndex As Integer) As Boolean
    ShowLabel = ((SelectedIndex + 1) And 1) <> 0
End Function
