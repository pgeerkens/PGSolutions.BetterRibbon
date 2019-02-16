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
Private Const COMAddInName            As String = "PGSolutions.BetterRibbon"
Private Const ControlStringsRangeName As String = "ControlStringData"

Private MBetterRibbon        As PGSolutions_RibbonDispatcher.IRibbonDispatcher
Private CustomGroup          As PGSolutions_RibbonDispatcher.RibbonGroupModel

Private WithEvents Toggle1   As PGSolutions_RibbonDispatcher.RibbonToggleModel
Attribute Toggle1.VB_VarHelpID = -1
Private WithEvents Toggle2   As PGSolutions_RibbonDispatcher.RibbonToggleModel
Attribute Toggle2.VB_VarHelpID = -1
Private WithEvents Toggle3   As PGSolutions_RibbonDispatcher.RibbonToggleModel
Attribute Toggle3.VB_VarHelpID = -1
Private WithEvents Toggle4   As PGSolutions_RibbonDispatcher.RibbonToggleModel
Attribute Toggle4.VB_VarHelpID = -1

Private WithEvents DropDown1 As PGSolutions_RibbonDispatcher.RibbonDropDownModel
Attribute DropDown1.VB_VarHelpID = -1
Private WithEvents DropDown2 As PGSolutions_RibbonDispatcher.RibbonDropDownModel
Attribute DropDown2.VB_VarHelpID = -1
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
    CustomGroup.SetShowInactive IsPressed
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
    DropDown1.IsEnabled = Toggle3.IsPressed
    
    Invalidate
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, TypeName(Me) & ".Toggle3_Toggled", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub

Private Sub Toggle4_Toggled(ByVal Sender As Variant, ByVal IsPressed As Boolean)
    On Error GoTo EH
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Control Strings")
    If ws.Visible = xlSheetVeryHidden Then
        ws.Visible = xlSheetVisible
    Else
        ws.Visible = xlSheetVeryHidden
    End If
    Toggle4.IsPressed = (ws.Visible = xlSheetVisible)
    ws.Activate
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, TypeName(Me) & ".Toggle3_Toggled", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub

Private Sub DropDown1_SelectionMade(ByVal Sender As Variant, ByVal SelectedIndex As Long)
    On Error GoTo EH
    SetImageAndLabel SelectedIndex, Button1, Button2, Button3
    
    Invalidate
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, TypeName(Me) & ".DropDown1_SelectionMade", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub

Private Sub DropDown2_SelectionMade(ByVal Sender As Variant, ByVal SelectedIndex As Long)
    On Error GoTo EH
    MsgBox "DropDown2 has not been implemented yet", vbOKOnly Or vbInformation, TypeName(Me)
    Invalidate
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, TypeName(Me) & ".DropDown2_SelectionMade", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub

Private Sub DropDown3_SelectionMade(ByVal Sender As Variant, ByVal SelectedIndex As Long)
    On Error GoTo EH
    MsgBox "DropDown2 has not been implemented yet", vbOKOnly Or vbInformation, TypeName(Me)
    Invalidate
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, TypeName(Me) & ".DropDown3_SelectionMade", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub

Private Sub Button1_Clicked(ByVal Sender As Variant)
    On Error GoTo EH
    Static IsToggle As Boolean
    IsToggle = AlternateToggle(BetterRibbon, IsToggle, Toggle2, "CustomVbaToggle2", "CustomVbaCheckBox2")
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, TypeName(Me) & ".Button1_Clicked", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub

Private Sub Button2_Clicked(ByVal Sender As Variant)
    On Error GoTo EH
    Static IsToggle As Boolean
    IsToggle = AlternateToggle(BetterRibbon, IsToggle, Toggle3, "CustomVbaToggle3", "CustomVbaCheckBox3")
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, TypeName(Me) & ".Button2_Clicked", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub

Private Sub Button3_Clicked(ByVal Sender As Variant)
    On Error GoTo EH
    If Not CustomGroup Is Nothing Then CustomGroup.Detach
    Class_Initialize
    Activate
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, TypeName(Me) & ".Button3_Clicked", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub

Private Sub Invalidate()
    On Error GoTo EH
    
    If Not Toggle1 Is Nothing Then Toggle1.Invalidate
    If Not Toggle2 Is Nothing Then Toggle2.Invalidate
    If Not Toggle3 Is Nothing Then Toggle3.Invalidate
    If Not Toggle4 Is Nothing Then Toggle4.Invalidate
    
    If Not DropDown1 Is Nothing Then DropDown1.Invalidate
    If Not DropDown2 Is Nothing Then DropDown2.Invalidate
    If Not Dropdown3 Is Nothing Then Dropdown3.Invalidate
    
    If Not Button1 Is Nothing Then Button1.Invalidate
    If Not Button2 Is Nothing Then Button2.Invalidate
    If Not Button3 Is Nothing Then Button3.Invalidate
    
    If Not CustomGroup Is Nothing Then CustomGroup.Invalidate
    Application.StatusBar = "Ready ...'"
XT: Exit Sub
EH: ErrorUtils.ReRaiseError Err, TypeName(Me) & ".Invalidate"
    Resume          ' for debugging only
End Sub

Friend Sub Activate()
    On Error GoTo EH
    If Not CustomGroup Is Nothing Then CustomGroup.Attach "CustomizableGroup"
    
    If Not Toggle1 Is Nothing Then Toggle1.Attach "CustomVbaToggle1"
    If Not Toggle2 Is Nothing Then Toggle2.Attach "CustomVbaCheckBox2"
    If Not Toggle3 Is Nothing Then Toggle3.Attach "CustomVbaCheckBox3"
    If Not Toggle4 Is Nothing Then Toggle4.Attach "CustomVbaCheckBox1"
    
    If Not DropDown1 Is Nothing Then DropDown1.Attach "CustomVbaDropDown1"
    If Not DropDown2 Is Nothing Then DropDown2.Attach "CustomVbaDropDown2"
    If Not Dropdown3 Is Nothing Then Dropdown3.Attach "CustomVbaDropDown3"
    
    If Not Button1 Is Nothing Then Button1.Attach "CustomizableButton1"
    If Not Button2 Is Nothing Then Button2.Attach "CustomizableButton2"
    If Not Button3 Is Nothing Then Button3.Attach "CustomizableButton3"
    
    ' Set initial status
    Toggle4.IsPressed = (ThisWorkbook.Worksheets("Control Strings").Visible = xlSheetVisible)
    If Not CustomGroup Is Nothing Then CustomGroup.SetShowInactive Toggle1.IsPressed
    DropDown1.IsEnabled = Toggle3.IsPressed
    
    Invalidate
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
    With BetterRibbon
        Set CustomGroup = .NewRibbonGroupModel(FindStrings("CustomizableGroup"))
    
        Set Toggle1 = .NewRibbonToggleModelMso(FindStrings("Toggle1"), "MarginsShowHide")
        Set Toggle2 = .NewRibbonToggleModelMso(FindStrings("Toggle2"), "MarginsShowHide")
        Set Toggle3 = .NewRibbonToggleModelMso(FindStrings("Toggle3"), "MarginsShowHide")
        Set Toggle4 = .NewRibbonToggleModelMso(FindStrings("Toggle4"), "MarginsShowHide")
        
        Set DropDown1 = .NewRibbonDropDownModel(FindStrings("DropDown1"))
        DropDown1.AddSelectableModel(.NewSelectableModel("LabelOnly", FindStrings("LabelOnly"))) _
                 .AddSelectableModel(.NewSelectableModel("ImageOnly", FindStrings("ImageOnly"))) _
                 .AddSelectableModel(.NewSelectableModel("LabelAndImage", FindStrings("LabelAndImage"))) _
                 .SelectedIndex = 2
        
        Set Button1 = .NewRibbonButtonModelMso(FindStrings("Button1"), "MacroPlay")
        Set Button2 = .NewRibbonButtonModelMso(FindStrings("Button2"), "MacroPlay")
        Set Button3 = .NewRibbonButtonModelMso(FindStrings("Button3"), "Refresh")
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

Private Function FindStrings(ByVal Search As String) As IRibbonControlStrings
    Const ControlStringsRangeName As String = "ControlStringData"
    
    On Error GoTo EH
    With New RibbonControlStrings
        Set FindStrings = .Initialize( _
                KeyTip:=FindStr(Search, Range(ControlStringsRangeName), 2), _
                Label:=FindStr(Search, Range(ControlStringsRangeName), 3), _
                AlternateLabel:=FindStr(Search, Range(ControlStringsRangeName), 4), _
                ScreenTip:=FindStr(Search, Range(ControlStringsRangeName), 5), _
                SuperTip:=FindStr(Search, Range(ControlStringsRangeName), 6), _
                Description:=FindStr(Search, Range(ControlStringsRangeName), 7))
    End With
XT: Exit Function
EH: ErrorUtils.ReRaiseError Err, TypeName(Me) & ".BetterRibbon"
    Resume          ' for debugging only
End Function

Private Function FindStr(ByVal Search As String, ByVal Range As Range, ByVal ColNo As Integer) As String
    On Error GoTo EH
    Dim v As Variant: v = Application.WorksheetFunction.VLookup(Search, Range, ColNo, False)
    FindStr = IIf(VarType(v) = vbString, v, "")
XT: Exit Function
EH: If Err.Number = 1004 Then Resume Next
    ErrorUtils.ReRaiseError Err, TypeName(Me) & ".BetterRibbon"
    Resume          ' for debugging only
End Function
