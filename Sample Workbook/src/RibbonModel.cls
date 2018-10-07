VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RibbonModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

#Const UseCheckBox = False

Private Strings                         As ResourceLoader

Private MViewModel                      As RibbonViewModel
Private RibbonTab                       As RibbonCommon

Private BrandingGroup                   As RibbonGroup
Attribute BrandingGroup.VB_VarHelpID = -1
Private WithEvents BrandingButton       As RibbonButton
Attribute BrandingButton.VB_VarHelpID = -1

Private FirstGroup                      As RibbonGroup
Attribute FirstGroup.VB_VarHelpID = -1
Private WithEvents StandardButton1      As RibbonButton
Attribute StandardButton1.VB_VarHelpID = -1
Private WithEvents StandardButton2      As RibbonButton
Attribute StandardButton2.VB_VarHelpID = -1
Private WithEvents StandardOptions      As RibbonDropDown
Attribute StandardOptions.VB_VarHelpID = -1
#If UseCheckBox Then
    Private WithEvents ShowAdvancedToggle  As RibbonCheckBox
Attribute ShowAdvancedToggle.VB_VarHelpID = -1
#Else
    Private WithEvents ShowAdvancedToggle  As RibbonToggleButton
Attribute ShowAdvancedToggle.VB_VarHelpID = -1
#End If

Private SecondGroup                     As RibbonGroup
Attribute SecondGroup.VB_VarHelpID = -1
Private WithEvents Group2Button1        As RibbonButton
Attribute Group2Button1.VB_VarHelpID = -1
Private WithEvents Group2Button2        As RibbonButton
Attribute Group2Button2.VB_VarHelpID = -1
Private WithEvents Group2Button3        As RibbonButton
Attribute Group2Button3.VB_VarHelpID = -1
Private WithEvents CustomSizeToggle     As RibbonToggleButton
Attribute CustomSizeToggle.VB_VarHelpID = -1
Private WithEvents Group2Options        As RibbonDropDown
Attribute Group2Options.VB_VarHelpID = -1

Private Sub BrandingButton_Clicked()
    MsgBox "Canadien, eh!", vbOKOnly Or vbInformation, "PGSolutions Ribbon Refresh"
End Sub

Private Sub Group2Button1_Clicked()
    DefaultButtonAction Group2Button1
End Sub
Private Sub Group2Button2_Clicked()
    DefaultButtonAction Group2Button2
End Sub
Private Sub Group2Button3_Clicked()
    DefaultButtonAction Group2Button3
End Sub
Private Sub Group2Options_SelectionMade(ByVal ItemId As String, ByVal ItemIndex As Long)
    SetButtonView ItemIndex, Group2Button1, Group2Button2, Group2Button3
End Sub

Private Sub CustomSizeToggle_Toggled(ByVal IsPressed As Boolean)
    Group2Options.IsEnabled = _
            ToggleCustomSize(Not IsPressed, Group2Button1, Group2Button2, Group2Button3)
End Sub

Private Sub StandardButton1_Clicked()
    DefaultButtonAction StandardButton1
End Sub
Private Sub StandardButton2_Clicked()
    RibbonLoader.ReInitializeRibbon
End Sub
Private Sub StandardOptions_SelectionMade(ByVal ItemId As String, ByVal ItemIndex As Long)
    SetButtonView ItemIndex, StandardButton1, StandardButton2
End Sub
Private Sub ShowAdvancedToggle_Toggled(ByVal IsPressed As Boolean)
    SecondGroup.IsVisible = IsPressed
End Sub

Friend Function Initialize(ByVal ribbonUI As IRibbonUI) As RibbonModel
    On Error GoTo EH
    Dim HideGrp2Undefined   As Boolean    ' When true hides the undefined controls in Group2
    HideGrp2Undefined = True
    Dim HideGrp1Undefined As Boolean
    HideGrp1Undefined = True
    
    Set MViewModel = AddInHandle.NewRibbonViewModel(ribbonUI)
    With MViewModel.RibbonFactory
        Dim Discard As Object
        
        Set Strings = .NewResourceLoader
        Set RibbonTab = .NewRibbonGroup( _
            Strings.AddControlStrings(ControlId:="TabPGSolutionsSampleWorkbook", _
                Label:="Sample App #1"))
    
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''' Branding Group
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Set BrandingGroup = .NewRibbonGroup( _
            Strings.AddControlStrings(ControlId:="BrandingGroup", _
                Label:="PGSolutions", _
                ScreenTip:="Tools for VBA Development"))
        Set BrandingButton = .NewRibbonButtonMso( _
            Strings.AddControlStrings(ControlId:="BrandingButton", _
                Label:="About Ribbon", _
                ScreenTip:="A Sample Pre-Plumbed Ribbon", _
                SuperTip:="An example of a pre-plumbed ribbon with improved error handling"), _
            ImageMso:="MacroSecurity", _
            ShowImage:=True)
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''' Group 1
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Set FirstGroup = .NewRibbonGroup( _
            Strings.AddControlStrings(ControlId:="Group1", _
                Label:="Standard Buttons", _
                ScreenTip:="Buttons Always Visible", _
                SuperTip:="These buttons are always visible on this ribbon."), _
            Visible:=True, Enabled:=True)
            
        If HideGrp1Undefined Then _
        Set Discard = .NewRibbonToggle( _
            Strings.AddControlStrings(ControlId:="Group1Toggle1", _
                Label:="Show Group #2"), _
            Visible:=False, Enabled:=False)
        Set StandardOptions = .NewRibbonDropDown( _
            Strings.AddControlStrings(ControlId:="Group1Dropdown1", _
                Label:="Options", _
                ScreenTip:="Controls Button Display", _
                SuperTip:="Display options for Standard Buttons:" & _
                        vbNewLine & "Label Only;" & _
                        vbNewLine & "Image Only;" & _
                        vbNewLine & "Label & Image"), _
            Visible:=True, Enabled:=True)
        StandardOptions.AddItem .NewSelectableItem( _
            Strings.AddControlStrings(ControlId:="Grp1Dropdown1_LabelOnly", _
                Label:="Label Only", _
                ScreenTip:="Display Label Only", _
                SuperTip:="Displays just the button label when its size is Regular; else both."))
        StandardOptions.AddItem .NewSelectableItem( _
            Strings.AddControlStrings(ControlId:="Grp1Dropdown1_ImageOnly", _
                Label:="Image Only", _
                ScreenTip:="Display Image Only", _
                SuperTip:="Displays just the button image when its size is Regular; else both."))
        StandardOptions.AddItem .NewSelectableItem( _
            Strings.AddControlStrings(ControlId:="Grp1Dropdown1_LabelAndImage", _
                Label:="Label & Image", _
                ScreenTip:="Display Label & Image", _
                SuperTip:="Always displays both the button label and image."))
        If HideGrp1Undefined Then _
        Set Discard = .NewRibbonDropDown( _
            Strings.AddControlStrings(ControlId:="Group1Dropdown2", Label:=""), _
            Visible:=False, Enabled:=False)
        If HideGrp1Undefined Then _
        Set Discard = .NewRibbonDropDown( _
            Strings.AddControlStrings(ControlId:="Group1Dropdown3", Label:=""), _
            Visible:=False, Enabled:=False)
        
        Set StandardButton1 = .NewRibbonButtonMso( _
            Strings.AddControlStrings(ControlId:="Group1Button1", _
                Label:="1st Std Button", _
                ScreenTip:="First Std Button", _
                SuperTip:="Demonstration of a 'standard button'; one that is always visible on the ribbon."), _
            Visible:=True, _
            Enabled:=True, _
            Size:=rdRegular, _
            ImageMso:="RefreshAll", _
            ShowImage:=True, _
            ShowLabel:=True)
        Set StandardButton2 = .NewRibbonButtonMso( _
            Strings.AddControlStrings(ControlId:="Group1Button2", _
                Label:="Refresh Ribbon", _
                ScreenTip:="Reinitializes Ribbon", _
                SuperTip:="Reinitializees the ribbon, as long as the original RibbonUI value is valid."), _
            Visible:=True, _
            Enabled:=True, _
            Size:=rdRegular, _
            ImageMso:="Refresh", _
            ShowImage:=True, _
            ShowLabel:=True)
        If HideGrp1Undefined Then _
        Set Discard = .NewRibbonButtonMso("Group1Button3", False, False, rdRegular)
        
        If HideGrp1Undefined Then _
        Set Discard = .NewRibbonCheckBox( _
            Strings.AddControlStrings(ControlId:="Group1Checkbox1", Label:=""), _
            Visible:=False, Enabled:=False)
        If HideGrp1Undefined Then _
        Set Discard = .NewRibbonCheckBox( _
            Strings.AddControlStrings(ControlId:="Group1Checkbox2", Label:=""), _
            Visible:=False, Enabled:=False)
            
#If UseCheckBox Then
        Set ShowAdvancedToggle = .NewRibbonCheckBox( _
            Strings.AddControlStrings(ControlId:="Group1Checkbox3", _
                Label:="Show Group #2", _
                ScreenTip:="Toggles Display of Group #2 Buttons", _
                SuperTip:="Alternates between setting the Group #2 buttons visible and hidden.", _
                AlternateLabel:=""), _
            Visible:=True, Enabled:=True)
        
        If HideGrp1Undefined Then _
        Set Discard = .NewRibbonToggle( _
            Strings.AddControlStrings(ControlId:="Group1Toggle2", Label:=""), _
            Visible:=False, Enabled:=False)
#Else
        If HideGrp1Undefined Then _
        Set Discard = .NewRibbonCheckBox( _
            Strings.AddControlStrings(ControlId:="Group1Checkbox3", Label:=""), _
            Visible:=False, Enabled:=False)
        
        Set ShowAdvancedToggle = .NewRibbonToggle( _
            Strings.AddControlStrings(ControlId:="Group1Toggle2", _
                Label:="Show Group #2", _
                ScreenTip:="Toggles Display of Group #2 Buttons", _
                SuperTip:="Alternates between setting the Group #2 buttons visible and hidden.", _
                AlternateLabel:=""), _
            Visible:=True, Enabled:=True)
#End If
                
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''' Group 2
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Set SecondGroup = .NewRibbonGroup( _
            Strings.AddControlStrings(ControlId:="Group2", _
                Label:="Group #2 Controls", _
                ScreenTip:="Buttons Optionally Visible", _
                SuperTip:="These buttons are visible on this ribbon only when requested."), _
            True)
        Set CustomSizeToggle = .NewRibbonToggleMso( _
            Strings.AddControlStrings(ControlId:="Group2Toggle1", _
                Label:="Prefer Small", _
                ScreenTip:="Toggles Button Size", _
                SuperTip:="Toggles between large and small (preferred size for) buttons for this ribbon group.", _
                AlternateLabel:="Prefer Large"), _
            True, True, rdLarge, DefaultImage, False, True)
        
        Set Group2Options = .NewRibbonDropDown( _
            Strings.AddControlStrings(ControlId:="Group2Dropdown1", _
                Label:="Display Options", _
                ScreenTip:="Group #2 Button Options", _
                SuperTip:="Display options for Group #2 Buttons:" & _
                        vbNewLine & "Label Only;" & _
                        vbNewLine & "Image Only;" & _
                        vbNewLine & "Label & Image"), _
            Visible:=True, Enabled:=False)
        Group2Options.AddItem .NewSelectableItem( _
            Strings.AddControlStrings(ControlId:="Group2Dropdown1_LabelOnly", _
                Label:="Label Only", _
                ScreenTip:="Display Label Only", _
                SuperTip:="Displays just the button label when its size is Regular; else both."))
        Group2Options.AddItem .NewSelectableItem( _
            Strings.AddControlStrings(ControlId:="Group2Dropdown1_ImageOnly", _
                Label:="Image Only", _
                ScreenTip:="Display Image Only", _
                SuperTip:="Displays just the button image when its size is Regular; else both."))
        Group2Options.AddItem .NewSelectableItem( _
            Strings.AddControlStrings(ControlId:="Group2Dropdown1_LabelAndImage", _
                Label:="Label & Image", _
                ScreenTip:="Display Label & Image", _
                SuperTip:="Always displays both the button label and image."))
If HideGrp2Undefined Then
        Set Discard = .NewRibbonDropDown( _
            Strings.AddControlStrings(ControlId:="Group2Dropdown2", Label:=""), _
            Visible:=False, Enabled:=False)
        Set Discard = .NewRibbonDropDown( _
            Strings.AddControlStrings(ControlId:="Group2Dropdown3", Label:=""), _
            Visible:=False, Enabled:=False)
End If
        
        Set Group2Button1 = .NewRibbonButtonMso( _
            Strings.AddControlStrings(ControlId:="Group2Button1", _
                Label:="Group #2 1st Button", _
                ScreenTip:="First Group #2 Button", _
                SuperTip:="First example of a custom button."), _
            Visible:=True, _
            Enabled:=True, _
            Size:=rdLarge, _
            ImageMso:="RefreshAll", _
            ShowImage:=True, _
            ShowLabel:=True)
        Set Group2Button2 = .NewRibbonButtonMso( _
            Strings.AddControlStrings(ControlId:="Group2Button2", _
                Label:="Group #2 2nd Button", _
                ScreenTip:="Seccond Group #2 Button", _
                SuperTip:="Second example of a custom button."), _
            Visible:=True, _
            Enabled:=True, _
            Size:=rdLarge, _
            ImageMso:="Refresh", _
            ShowImage:=True, _
            ShowLabel:=True)
        Set Group2Button3 = .NewRibbonButtonMso( _
            Strings.AddControlStrings(ControlId:="Group2Button3", _
                Label:="Group #2 3rd Button", _
                ScreenTip:="Third Group #2 Button", _
                SuperTip:="Third example of a custom button."), _
            Visible:=True, _
            Enabled:=True, _
            Size:=rdLarge, _
            ImageMso:="MacroPlay", _
            ShowImage:=True, _
            ShowLabel:=True)
        
If HideGrp2Undefined Then
        Set Discard = .NewRibbonCheckBox( _
            Strings.AddControlStrings(ControlId:="Group2Checkbox1", Label:=""), _
            Visible:=False, Enabled:=False)
        Set Discard = .NewRibbonCheckBox( _
            Strings.AddControlStrings(ControlId:="Group2Checkbox2", Label:=""), _
            Visible:=False, Enabled:=False)
        Set Discard = .NewRibbonCheckBox( _
            Strings.AddControlStrings(ControlId:="Group2Checkbox3", Label:=""), _
            Visible:=False, Enabled:=False)
        
        Set Discard = .NewRibbonToggle( _
            Strings.AddControlStrings(ControlId:="Group2Toggle2", Label:=""), _
            Visible:=False, Enabled:=False)
End If

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''' Default Settings
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ShowAdvancedToggle.OnToggled False
        StandardOptions.OnActionDropDown "", 2
    
        Group2Options.SelectedItemId = "Group2Dropdown1_LabelAndImage"
        Group2Options.IsEnabled = CustomSizeToggle.IsPressed
    End With
    MViewModel.Invalidate
    
    Set Initialize = Me
XT: Exit Function
EH: ReraiseError Err, TypeName(Me) & ".Initialize"
    Resume          ' for debugging only
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''   DON'T EDIT THESE ROUTINES   '''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Property Get ViewModel() As RibbonViewModel
    Set ViewModel = MViewModel
End Property

Friend Sub ActivateTab()
    MViewModel.ActivateTab RibbonTab.ID
End Sub
