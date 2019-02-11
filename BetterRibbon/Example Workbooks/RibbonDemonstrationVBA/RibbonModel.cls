VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RibbonModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Button1Strings      As PGSolutions_RibbonDispatcher.IRibbonControlStrings
Private Button2Strings      As PGSolutions_RibbonDispatcher.IRibbonControlStrings
Private Button3Strings      As PGSolutions_RibbonDispatcher.IRibbonControlStrings
Private Toggle1Strings      As PGSolutions_RibbonDispatcher.IRibbonControlStrings
Private Toggle2Strings      As PGSolutions_RibbonDispatcher.IRibbonControlStrings
Private Toggle3Strings      As PGSolutions_RibbonDispatcher.IRibbonControlStrings

Private WithEvents Button1  As PGSolutions_RibbonDispatcher.RibbonButton
Attribute Button1.VB_VarHelpID = -1
Private WithEvents Button2  As PGSolutions_RibbonDispatcher.RibbonButton
Attribute Button2.VB_VarHelpID = -1
Private WithEvents Button3  As PGSolutions_RibbonDispatcher.RibbonButton
Attribute Button3.VB_VarHelpID = -1
Private WithEvents Toggle1  As PGSolutions_RibbonDispatcher.RibbonToggleModel
Attribute Toggle1.VB_VarHelpID = -1
Private WithEvents Toggle2  As PGSolutions_RibbonDispatcher.RibbonToggleModel
Attribute Toggle2.VB_VarHelpID = -1
Private WithEvents Toggle3  As PGSolutions_RibbonDispatcher.RibbonToggleModel
Attribute Toggle3.VB_VarHelpID = -1

Private Sub Button1_Clicked(ByVal Sender As Variant)
    On Error GoTo EH
    ButtonProcessing.Button1_Processing ThisWorkbook.Name
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, TypeName(Me) & ".Button1_Clicked", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub

Private Sub Button2_Clicked(ByVal Sender As Variant)
    On Error GoTo EH
    ButtonProcessing.Button2_Processing ThisWorkbook.Name
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, TypeName(Me) & ".Button2_Clicked", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub

Private Sub Button3_Clicked(ByVal Sender As Variant)
    On Error GoTo EH
    ButtonProcessing.Button3_Processing ThisWorkbook.Name
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, TypeName(Me) & ".Button3_Clicked", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub

Private Sub Toggle1_Toggled(ByVal Sender As Variant, ByVal IsPressed As Boolean)
    On Error GoTo EH
    RibbonDispatcher.ShowInactive IsPressed
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, TypeName(Me) & ".Toggle1_Toggled", vbOKOnly Or vbInformation
    Resume          ' for debugging only
End Sub

Private Sub Toggle2_Toggled(ByVal Sender As Variant, ByVal IsPressed As Boolean)
    On Error GoTo EH
    Toggle1.ViewModel.IsLarge = IsPressed
    Toggle2.ViewModel.IsLarge = IsPressed
    'Toggle3.ViewModel.IsLarge = IsPressed
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, TypeName(Me) & ".Toggle2_Toggled", vbOKOnly Or vbInformation
    Resume          ' for debugging only
End Sub

Private Sub Toggle3_Toggled(ByVal Sender As Variant, ByVal IsPressed As Boolean)
    On Error GoTo EH
    RibbonDispatcher.ShowInactive IsPressed
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, TypeName(Me) & ".Toggle3_Toggled", vbOKOnly Or vbInformation
    Resume          ' for debugging only
End Sub

Friend Sub Activate()
    On Error GoTo EH
    With RibbonDispatcher
        Set Button1 = .AttachButton("CustomizableButton1", Button1Strings)
        Button1.SetImageMso "RefreshAll"
        Set Button2 = .AttachButton("CustomizableButton2", Button2Strings)
        Button2.SetImageMso "Refresh"
        Set Button3 = .AttachButton("CustomizableButton3", Button3Strings)
        Button3.SetImageMso "MacroPlay"
    End With
    
    'If Toggle1 Is Nothing Then Set Toggle1 = RibbonDispatcher.NewRibbonToggleModel
    Toggle1.Attach "CustomVbaToggle1", Toggle1Strings
    
    'If Toggle2 Is Nothing Then Set Toggle2 = RibbonDispatcher.NewRibbonToggleModel
    Toggle2.Attach "CustomVbaToggle2", Toggle2Strings
    
    'If Toggle3 Is Nothing Then Set Toggle3 = RibbonDispatcher.NewRibbonToggleModel
    'Toggle3.Attach "CustomVbaToggle2", Toggle3Strings
    
    
    RibbonDispatcher.InvalidateCustomControlsGroup
XT: Exit Sub
EH: ErrorUtils.ReRaiseError Err, TypeName(Me) & ".Activate"
    Resume          ' for debugging only
End Sub

Private Sub Class_Initialize()
    On Error GoTo EH
    With RibbonDispatcher
        Set Button1Strings = .NewControlStrings(Label:="This is cool!", _
                ScreenTip:="Button1 Screentip", _
                SuperTip:="Lots of good things" & vbNewLine & _
                          "can be done here to" & vbNewLine & _
                          "show off a bit.", KeyTip:="", AlternateLabel:="", Description:="")
        Set Button2Strings = .NewControlStrings(Label:="This is cooler!", _
                ScreenTip:="Button2 Screentip", _
                SuperTip:="Lots of good things" & vbNewLine & _
                          "can be done from hither" & vbNewLine & _
                          " " & vbNewLine & _
                          " " & vbNewLine & _
                          " " & vbNewLine & _
                          " " & vbNewLine & _
                          " " & vbNewLine & _
                          " " & vbNewLine & _
                          "... all the way to yon." & vbNewLine & _
                          "to show off a bit more.")
        Set Button3Strings = .NewControlStrings(Label:="This is coolest!", _
                ScreenTip:="Button3 Screentip", _
                SuperTip:="What's the weather like where you are?.")
        
        Set Toggle1Strings = .NewControlStrings(Label:="Show Inactive", _
                ScreenTip:="Toggles Display of Inactive", _
                SuperTip:="Toggles on/off the display of customizable ribbon controls" & _
                          " that are currently unattached to a Data Source and/or" & _
                          "event sink.")
        Set Toggle2Strings = .NewControlStrings(Label:="Prefer Large", _
                ScreenTip:="Toggles Large or Small Toggles", _
                SuperTip:="Toggles the size of the attached Toggle" & _
                          " Buttons between Large and Regular.")
        'Set Toggle3Strings = .NewControlStrings(Label:="Show Inactive", _
                ScreenTip:="Toggles Display of Inactive", _
                SuperTip:="Toggles on/off the display of customizable ribbon controls" & _
                          " that are currently unattached to a Data Source and/or" & _
                          "event sink.")
                          
        Set Toggle1 = .NewRibbonToggleModel
        Set Toggle2 = .NewRibbonToggleModel
        Set Toggle3 = .NewRibbonToggleModel
        
        .Invalidate
    End With
XT: Exit Sub
EH: ErrorUtils.ReRaiseError Err, TypeName(Me) & ".Class_Initialize"
    Resume          ' for debugging only
End Sub
