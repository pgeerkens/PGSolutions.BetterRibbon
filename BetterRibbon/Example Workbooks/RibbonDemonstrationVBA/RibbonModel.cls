VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RibbonModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''=======================================================================================
'''                Copyright (c) 2017-2019 Pieter Geerkens
'''
'''     Licensed under the MIT Licence at:
'''             https://github.com/pgeerkens/PGSolutions.BetterRibbon/blob/dev/LICENSE
'''=======================================================================================
Option Explicit
Private Const COMAddInName   As String = "PGSolutions.BetterRibbon"

Private MBetterRibbon        As PGSolutions_RibbonDispatcher.IRibbonDispatcher
Private CustomGroup          As PGSolutions_RibbonDispatcher.RibbonGroup

Private WithEvents Toggle1   As PGSolutions_RibbonDispatcher.RibbonToggleModel
Attribute Toggle1.VB_VarHelpID = -1
Private WithEvents Toggle2   As PGSolutions_RibbonDispatcher.RibbonToggleModel
Attribute Toggle2.VB_VarHelpID = -1
Private WithEvents Toggle3   As PGSolutions_RibbonDispatcher.RibbonToggleModel
Attribute Toggle3.VB_VarHelpID = -1

Private WithEvents Dropdown1 As PGSolutions_RibbonDispatcher.RibbonDropDownModel
Attribute Dropdown1.VB_VarHelpID = -1
Private WithEvents Dropdown2 As PGSolutions_RibbonDispatcher.RibbonDropDownModel
Attribute Dropdown2.VB_VarHelpID = -1
Private WithEvents Dropdown3 As PGSolutions_RibbonDispatcher.RibbonDropDownModel
Attribute Dropdown3.VB_VarHelpID = -1

Private WithEvents Button1   As PGSolutions_RibbonDispatcher.RibbonButtonModel
Attribute Button1.VB_VarHelpID = -1
Private WithEvents Button2   As PGSolutions_RibbonDispatcher.RibbonButtonModel
Attribute Button2.VB_VarHelpID = -1
Private WithEvents Button3   As PGSolutions_RibbonDispatcher.RibbonButtonModel
Attribute Button3.VB_VarHelpID = -1

Private Sub Toggle1_Toggled(ByVal Sender As Variant, ByVal IsPressed As Boolean)
    On Error GoTo EH
    BetterRibbon.ShowInactive IsPressed
    Toggle1.SetImageMso ToggleImage(IsPressed)
    
    Invalidate
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
    
    Invalidate
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
    
    Invalidate
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, TypeName(Me) & ".Toggle3_Toggled", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub

Private Sub Dropdown1_SelectionMade(ByVal Sender As Variant, ByVal SelectedIndex As Long)
    On Error GoTo EH
    SetImageAndLabel SelectedIndex, Button1, Button2, Button3
    
    Invalidate
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, TypeName(Me) & ".Dropdown1_SelectionMade", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub

Private Sub Dropdown2_SelectionMade(ByVal Sender As Variant, ByVal SelectedIndex As Long)
    On Error GoTo EH
    MsgBox "Dropdown2 has not been implemented yet", vbOKOnly Or vbInformation, TypeName(Me)
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, TypeName(Me) & ".Dropdown2_SelectionMade", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub

Private Sub Dropdown3_SelectionMade(ByVal Sender As Variant, ByVal SelectedIndex As Long)
    On Error GoTo EH
    MsgBox "Dropdown2 has not been implemented yet", vbOKOnly Or vbInformation, TypeName(Me)
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, TypeName(Me) & ".Dropdown3_SelectionMade", vbOKOnly Or vbInformation
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
    If Not Toggle1 Is Nothing Then Toggle1.Attach "CustomVbaToggle1"
    If Not Toggle2 Is Nothing Then Toggle2.Attach "CustomVbaCheckBox2"
    If Not Toggle3 Is Nothing Then Toggle3.Attach "CustomVbaCheckBox3"
    
    If Not Dropdown1 Is Nothing Then Dropdown1.Attach "CustomVbaDropdown1"
    If Not Dropdown2 Is Nothing Then Dropdown2.Attach "CustomVbaDropdown2"
    If Not Dropdown3 Is Nothing Then Dropdown3.Attach "CustomVbaDropdown3"
    
    If Not Button1 Is Nothing Then Button1.Attach "CustomizableButton1"
    If Not Button2 Is Nothing Then Button2.Attach "CustomizableButton2"
    If Not Button3 Is Nothing Then Button3.Attach "CustomizableButton3"
    
    Invalidate
XT: Exit Sub
EH: Select Case MsgBoxAbortRetryIgnore(Err, TypeName(Me) & ".Activate")
        Case vbAbort:  ErrorUtils.ReRaiseError Err, TypeName(Me) & ".Activate"
        Case vbRetry:  Resume
        Case vbIgnore: Resume Next
    End Select
    Resume          ' for debugging only
End Sub

Private Sub Invalidate()
    On Error GoTo EH
    If Not Toggle1 Is Nothing Then Toggle1.Invalidate
    If Not Toggle2 Is Nothing Then Toggle2.Invalidate
    If Not Toggle3 Is Nothing Then Toggle3.Invalidate
    
    If Not Dropdown1 Is Nothing Then Dropdown1.Invalidate
    If Not Dropdown2 Is Nothing Then Dropdown2.Invalidate
    If Not Dropdown3 Is Nothing Then Dropdown3.Invalidate
    
    If Not Button1 Is Nothing Then Button1.Invalidate
    If Not Button2 Is Nothing Then Button2.Invalidate
    If Not Button3 Is Nothing Then Button3.Invalidate
    
    BetterRibbon.InvalidateControl "CustomizableGroup"
XT: Exit Sub
EH: ErrorUtils.ReRaiseError Err, TypeName(Me) & ".Invalidate"
    Resume          ' for debugging only
End Sub

Private Sub Class_Initialize()
    On Error GoTo EH
    With BetterRibbon
        Set Toggle1 = .NewRibbonToggleModelMso(Toggle1Strings, "MarginsShowHide")
        Set Toggle2 = .NewRibbonToggleModelMso(Toggle2Strings, "MarginsShowHide")
        Set Toggle3 = .NewRibbonToggleModelMso(Toggle3Strings, "MarginsShowHide")
        
        Set Dropdown1 = .NewRibbonDropdownModel(Dropdown1Strings)
        Dropdown1.AddItem(.NewSelectableItem("LabelOnly", "Label Only")) _
                 .AddItem(.NewSelectableItem("ImageOnly", "Image Only")) _
                 .AddItem(.NewSelectableItem("LabelAndImage", "Label and Image")) _
                 .SelectedIndex = 2
        
        Set Button1 = .NewRibbonButtonModelMso(Button1Strings, "RefreshAll")
        Set Button2 = .NewRibbonButtonModelMso(Button2Strings, "Refresh")
        Set Button3 = .NewRibbonButtonModelMso(Button3Strings, "MacroPlay")
    End With
XT: Exit Sub
EH: Select Case MsgBoxAbortRetryIgnore(Err, TypeName(Me) & ".Class_Initialize")
        Case vbAbort:  ErrorUtils.ReRaiseError Err, TypeName(Me) & ".Class_Initialize"
        Case vbRetry:  Resume
        Case vbIgnore: Resume Next
    End Select
    Resume          ' for debugging only
End Sub

Private Property Get BetterRibbon() As PGSolutions_RibbonDispatcher.IRibbonDispatcher
    On Error GoTo EH
    If MBetterRibbon Is Nothing Then
        Set MBetterRibbon = Application.COMAddIns(COMAddInName).Object.NewBetterRibbon()
    End If
    Set BetterRibbon = MBetterRibbon
XT: Exit Property
EH: ErrorUtils.ReRaiseError Err, TypeName(Me) & ".BetterRibbon"
    Resume          ' for debugging only
End Property
