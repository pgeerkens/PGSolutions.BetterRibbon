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

Private Const COMAddInName  As String = "PGSolutions.BetterRibbon"
Private MModelServer        As PGSolutions_RibbonDispatcher.IModelServer

Public Sub Register()
    On Error GoTo EH
    Application.COMAddIns(COMAddInName).Object.RegisterWorkbook ThisWorkbook.Name
    
XT: Exit Sub
EH: ErrorUtils.ReRaiseError Err, ModuleName & ".AlternateToggle"
    Resume          ' for debugging only
End Sub

Public Function AlternateToggle(ByVal Factory As IModelServer, Mode As Boolean, _
        Model As ToggleModel, ByVal ToggleID As String, ByVal CheckBoxID As String _
) As Boolean
    On Error GoTo EH
    AlternateToggle = Not Mode
    
    Model.Detach
    Model.Attach IIf(AlternateToggle, ToggleID, CheckBoxID)
    Model.SetImage ModelServer.NewImageObjectMso(ToggleImage(Model.IsPressed))
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

Public Property Get ModelServer() As PGSolutions_RibbonDispatcher.IModelServer
    On Error GoTo EH
    If MModelServer Is Nothing Then
        Set MModelServer = Application.COMAddIns(COMAddInName).Object _
            .NewBetterRibbon(New ResourceLoader)
    End If
    Set ModelServer = MModelServer
XT: Exit Property
EH: ErrorUtils.ReRaiseError Err, ModuleName & ".BetterRibbon"
    Resume          ' for debugging only
End Property
