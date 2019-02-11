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

Private Toggle1Strings       As PGSolutions_RibbonDispatcher.IRibbonControlStrings
Private Toggle2Strings       As PGSolutions_RibbonDispatcher.IRibbonControlStrings
Private Toggle3Strings       As PGSolutions_RibbonDispatcher.IRibbonControlStrings
Private Dropdown1Strings     As PGSolutions_RibbonDispatcher.IRibbonControlStrings
'Private Dropdown1Strings     As PGSolutions_RibbonDispatcher.IRibbonControlStrings
'Private Dropdown1Strings     As PGSolutions_RibbonDispatcher.IRibbonControlStrings
Private Button1Strings       As PGSolutions_RibbonDispatcher.IRibbonControlStrings
Private Button2Strings       As PGSolutions_RibbonDispatcher.IRibbonControlStrings
Private Button3Strings       As PGSolutions_RibbonDispatcher.IRibbonControlStrings

Private WithEvents Toggle1   As PGSolutions_RibbonDispatcher.RibbonToggleModel
Attribute Toggle1.VB_VarHelpID = -1
Private WithEvents Toggle2   As PGSolutions_RibbonDispatcher.RibbonToggleModel
Attribute Toggle2.VB_VarHelpID = -1
Private WithEvents Toggle3   As PGSolutions_RibbonDispatcher.RibbonToggleModel
Attribute Toggle3.VB_VarHelpID = -1
Private WithEvents Dropdown1 As PGSolutions_RibbonDispatcher.RibbonDropDownModel
Attribute Dropdown1.VB_VarHelpID = -1
'Private WithEvents Dropdown2 As PGSolutions_RibbonDispatcher.RibbonDropDownModel
'Private WithEvents Dropdown3 As PGSolutions_RibbonDispatcher.RibbonDropDownModel
Private WithEvents Button1   As PGSolutions_RibbonDispatcher.RibbonButtonModel
Attribute Button1.VB_VarHelpID = -1
Private WithEvents Button2   As PGSolutions_RibbonDispatcher.RibbonButtonModel
Attribute Button2.VB_VarHelpID = -1
Private WithEvents Button3   As PGSolutions_RibbonDispatcher.RibbonButtonModel
Attribute Button3.VB_VarHelpID = -1

Private Sub Toggle1_Toggled(ByVal Sender As Variant, ByVal IsPressed As Boolean)
    On Error GoTo EH
    RibbonDispatcher.ShowInactive IsPressed
    Toggle1.SetImageMso ToggleImage(IsPressed)
    Toggle1.Invalidate
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, TypeName(Me) & ".Toggle1_Toggled", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub

Private Sub Toggle2_Toggled(ByVal Sender As Variant, ByVal IsPressed As Boolean)
    On Error GoTo EH
    Toggle1.IsLarge = Not IsPressed
    Toggle2.IsLarge = Not IsPressed
    Toggle3.IsLarge = Not IsPressed
    Toggle2.SetImageMso ToggleImage(IsPressed)
    Toggle1.Invalidate
    Toggle2.Invalidate
    Toggle3.Invalidate
    RibbonDispatcher.InvalidateCustomControlsGroup
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, TypeName(Me) & ".Toggle2_Toggled", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub

Private Sub Toggle3_Toggled(ByVal Sender As Variant, ByVal IsPressed As Boolean)
    On Error GoTo EH
    Button1.IsLarge = Not IsPressed
    Button2.IsLarge = Not IsPressed
    Button3.IsLarge = Not IsPressed
    Toggle3.SetImageMso ToggleImage(IsPressed)
    Dropdown1.IsEnabled = IsPressed
    Button1.Invalidate
    Button2.Invalidate
    Button3.Invalidate
    Dropdown1.Invalidate
    RibbonDispatcher.InvalidateCustomControlsGroup
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, TypeName(Me) & ".Toggle3_Toggled", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub

Private Sub Dropdown1_SelectionMade(ByVal Sender As Variant, ByVal SelectedIndex As Long)
    On Error GoTo EH
    SetImageAndLabel SelectedIndex, Button1, Button2, Button3
    RibbonDispatcher.InvalidateCustomControlsGroup
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, TypeName(Me) & ".Toggle3_Toggled", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub

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

Friend Sub Activate()
    On Error GoTo EH
    Dim UseCheckbox As Integer: UseCheckbox = 1
    
    Toggle1.Attach "CustomVbaToggle1", Toggle1Strings
    Toggle2.Attach IIf(UseCheckbox, "CustomVbaCheckBox2", "CustomVbaToggle2"), Toggle2Strings
    Toggle3.Attach IIf(UseCheckbox, "CustomVbaCheckBox3", "CustomVbaToggle3"), Toggle3Strings
    
    Dropdown1.Attach "CustomVbaDropdown1", Dropdown1Strings
    
    Button1.Attach "CustomizableButton1", Button1Strings:  Button1.Invalidate
    Button2.Attach "CustomizableButton2", Button2Strings:  Button2.Invalidate
    Button3.Attach "CustomizableButton3", Button3Strings:  Button3.Invalidate
    
    RibbonDispatcher.InvalidateCustomControlsGroup
XT: Exit Sub
EH: Select Case MsgBoxAbortRetryIgnore(Err, TypeName(Me) & ".Activate")
        Case vbAbort:  ErrorUtils.ReRaiseError Err, TypeName(Me) & ".Activate"
        Case vbRetry:  Resume
        Case vbIgnore: Resume Next
    End Select
    Resume          ' for debugging only
End Sub

Private Sub Class_Initialize()
    On Error GoTo EH
    With RibbonDispatcher
        Set Toggle1Strings = .NewControlStrings(Label:="Show Inactive", _
                ScreenTip:="Toggles Display of Inactive", _
                SuperTip:="Toggles on/off the display of customizable ribbon controls" & _
                          " that are currently unattached to a Data Source and/or" & _
                          " event sink.")
        Set Toggle2Strings = .NewControlStrings(Label:="Prefer Small Toggles", _
                ScreenTip:="Toggles Large or Small Toggles", _
                SuperTip:="Toggles the size of the activated Toggle" & _
                          " Buttons between Large and Regular.")
        Set Toggle3Strings = .NewControlStrings(Label:="Prefer Small Buttons", _
                ScreenTip:="Toggles Large or Small Buttons", _
                SuperTip:="Toggles the size of the activated Action" & _
                          " Buttons between Large and Regular.")
                          
        Set Toggle1 = .NewRibbonToggleModel
        Set Toggle2 = .NewRibbonToggleModel
        Set Toggle3 = .NewRibbonToggleModel
        
        Set Dropdown1Strings = .NewControlStrings(Label:="Image or Label", _
                ScreenTip:="Select Image, Label, or Both", _
                SuperTip:="Controls display of Image, Label, or Both" & _
                          " for the customizable action buttons.")
        
        Set Dropdown1 = .NewRibbonDropdownModel
        Dropdown1.AddItem(.NewSelectableItem("LabelOnly", "Label Only")) _
                 .AddItem(.NewSelectableItem("ImageOnly", "Image Only")) _
                 .AddItem .NewSelectableItem("LabelAndImage", "Label and Image")
        Dropdown1.SelectedIndex = 2
        
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
        
        Set Button1 = .NewRibbonButtonModel:  Button1.SetImageMso "RefreshAll": Button1.Invalidate
        Set Button2 = .NewRibbonButtonModel:  Button2.SetImageMso "Refresh":    Button2.Invalidate
        Set Button3 = .NewRibbonButtonModel:  Button3.SetImageMso "MacroPlay":  Button3.Invalidate
            
        .Invalidate
    End With
XT: Exit Sub
EH: Select Case MsgBoxAbortRetryIgnore(Err, TypeName(Me) & ".Class_Initialize")
        Case vbAbort:  ErrorUtils.ReRaiseError Err, TypeName(Me) & ".Class_Initialize"
        Case vbRetry:  Resume
        Case vbIgnore: Resume Next
    End Select
    Resume          ' for debugging only
End Sub
