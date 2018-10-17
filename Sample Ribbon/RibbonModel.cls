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

#Const UseCheckBox = True

Private Type TRibbonModel
    ViewModel       As RibbonViewModel
    RibbonTab       As RibbonCommon
End Type
Private This As TRibbonModel

Private Strings                         As PGSolutions_RibbonDispatcher.ControlStrings

Private BrandingGroup                   As RibbonGroup
Attribute BrandingGroup.VB_VarHelpID = -1
Private WithEvents BrandingButton       As RibbonButton
Attribute BrandingButton.VB_VarHelpID = -1

Private StandardGroup                   As RibbonGroup
Attribute StandardGroup.VB_VarHelpID = -1
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

Private CustomGroup                     As RibbonGroup
Attribute CustomGroup.VB_VarHelpID = -1
Private WithEvents AdvancedButton1      As RibbonButton
Attribute AdvancedButton1.VB_VarHelpID = -1
Private WithEvents AdvancedButton2      As RibbonButton
Attribute AdvancedButton2.VB_VarHelpID = -1
Private WithEvents AdvancedButton3      As RibbonButton
Attribute AdvancedButton3.VB_VarHelpID = -1
Private WithEvents CustomSizeToggle     As RibbonToggleButton
Attribute CustomSizeToggle.VB_VarHelpID = -1
Private WithEvents AdvancedOptions      As RibbonDropDown
Attribute AdvancedOptions.VB_VarHelpID = -1

Implements PGSolutions_RibbonDispatcher.IResourceManager

Private Sub BrandingButton_Clicked()
    MsgBox "Canadien, eh!", vbOKOnly Or vbInformation, "PGSolutions Ribbon Refresh"
End Sub

Private Sub AdvancedButton1_Clicked()
    DefaultButtonAction AdvancedButton1
End Sub
Private Sub AdvancedButton2_Clicked()
    DefaultButtonAction AdvancedButton2
End Sub
Private Sub AdvancedButton3_Clicked()
    DefaultButtonAction AdvancedButton3
End Sub
Private Sub AdvancedOptions_SelectionMade(ByVal ItemId As String, ByVal ItemIndex As Long)
    SetButtonView ItemIndex, AdvancedButton1, AdvancedButton2, AdvancedButton3
End Sub

Private Sub CustomSizeToggle_Toggled(ByVal IsPressed As Boolean)
    AdvancedOptions.IsEnabled = _
            ToggleCustomSize(Not IsPressed, AdvancedButton1, AdvancedButton2, AdvancedButton3)
End Sub

Private Sub StandardButton1_Clicked()
    DefaultButtonAction StandardButton1
End Sub
Private Sub StandardButton2_Clicked()
    DefaultButtonAction StandardButton2
End Sub
Private Sub StandardOptions_SelectionMade(ByVal ItemId As String, ByVal ItemIndex As Long)
    SetButtonView ItemIndex, StandardButton1, StandardButton2
End Sub
Private Sub ShowAdvancedToggle_Toggled(ByVal IsPressed As Boolean)
    CustomGroup.IsVisible = IsPressed
End Sub

Public Function Initialize(ByVal RibbonUI As IRibbonUI) As RibbonModel
    On Error GoTo EH
    Set Strings = New PGSolutions_RibbonDispatcher.ControlStrings
    
    Set This.ViewModel = NewRibbonViewModel(RibbonUI, Me)
    With This.ViewModel.RibbonFactory
        Dim Discard As Object
        
        Set This.RibbonTab = .NewRibbonGroup( _
            Strings.AddControl(ControlID:="TabPGSolutions", _
                Label:="Sample App #1"))
    
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''' Branding Group
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Set BrandingGroup = .NewRibbonGroup( _
            Strings.AddControl(ControlID:="BrandingGroup", _
                Label:="PGSolutions", _
                ScreenTip:="Tools for VBA Development"))
        Set BrandingButton = .NewRibbonButtonMso( _
            Strings.AddControl(ControlID:="BrandingButton", _
                Label:="About Ribbon", _
                ScreenTip:="A Sample Pre-Plumbed Ribbon", _
                SuperTip:="An example of a pre-plumbed ribbon with improved error handling"), _
            ImageMso:="MacroSecurity", _
            ShowImage:=True)
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''' Group 1
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Set StandardGroup = .NewRibbonGroup( _
            Strings.AddControl(ControlID:="Group1", _
                Label:="Standard Buttons", _
                ScreenTip:="Buttons Always Visible", _
                SuperTip:="These buttons are always visible on this ribbon."), _
            Visible:=True, Enabled:=True)
            
        Set StandardOptions = .NewRibbonDropDown( _
            Strings.AddControl(ControlID:="Group1Dropdown1", _
                Label:="Options", _
                ScreenTip:="Controls Button Display", _
                SuperTip:="Display options for Standard Buttons:" & _
                        vbNewLine & "Label Only;" & _
                        vbNewLine & "Image Only;" & _
                        vbNewLine & "Label & Image"), _
            Visible:=True, Enabled:=True)
        StandardOptions.AddItem .NewSelectableItem( _
            Strings.AddControl(ControlID:="Grp1Dropdown1_LabelOnly", _
                Label:="Label Only", _
                ScreenTip:="Display Label Only", _
                SuperTip:="Displays just the button label when its size is Regular; else both."))
        StandardOptions.AddItem .NewSelectableItem( _
            Strings.AddControl(ControlID:="Grp1Dropdown1_ImageOnly", _
                Label:="Image Only", _
                ScreenTip:="Display Image Only", _
                SuperTip:="Displays just the button image when its size is Regular; else both."))
        StandardOptions.AddItem .NewSelectableItem( _
            Strings.AddControl(ControlID:="Grp1Dropdown1_LabelAndImage", _
                Label:="Label & Image", _
                ScreenTip:="Display Label & Image", _
                SuperTip:="Always displays both the button label and image."))
        Set Discard = .NewRibbonDropDown( _
            Strings.AddControl(ControlID:="Group1Dropdown2", Label:=""), _
            Visible:=False, Enabled:=False)
        Set Discard = .NewRibbonDropDown( _
            Strings.AddControl(ControlID:="Group1Dropdown3", Label:=""), _
            Visible:=False, Enabled:=False)
        
        Set StandardButton1 = .NewRibbonButtonMso( _
            Strings.AddControl(ControlID:="Group1Button1", _
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
            Strings.AddControl(ControlID:="Group1Button2", _
                Label:="2nd Std Button", _
                ScreenTip:="Second Std Button", _
                SuperTip:="Another demonstration of a 'standard button'; one that is always visible on the ribbon."), _
            Visible:=True, _
            Enabled:=True, _
            Size:=rdRegular, _
            ImageMso:="Refresh", _
            ShowImage:=True, _
            ShowLabel:=True)
        Set Discard = .NewRibbonButtonMso("Group1Button3", False, False, rdRegular)
        
        Set Discard = .NewRibbonCheckBox( _
            Strings.AddControl(ControlID:="Group1Checkbox1", Label:=""), _
            Visible:=False, Enabled:=False)
        Set Discard = .NewRibbonCheckBox( _
            Strings.AddControl(ControlID:="Group1Checkbox2", Label:=""), _
            Visible:=False, Enabled:=False)
            
#If UseCheckBox Then
        Set ShowAdvancedToggle = .NewRibbonCheckBox( _
            Strings.AddControl(ControlID:="Group1Checkbox3", _
                Label:="Show Advanced", _
                ScreenTip:="Toggles Display of Advanced Buttons", _
                SuperTip:="Alternates between setting the advanced buttons visible and hidden.", _
                AlternateLabel:=""), _
            Visible:=True, Enabled:=True)
        
        Set Discard = .NewRibbonToggle( _
            Strings.AddControl(ControlID:="Group1Toggle1", Label:=""), _
            Visible:=False, Enabled:=False)
#Else
        Set Discard = .NewRibbonCheckBox( _
            Strings.AddControl(ControlID:="Group1Checkbox3", Label:=""), _
            Visible:=False, Enabled:=False)
        
        Set ShowAdvancedToggle = .NewRibbonToggle( _
            Strings.AddControl(ControlID:="Group1Toggle1", _
                Label:="Show Advanced", _
                ScreenTip:="Toggles Display of Advanced Buttons", _
                SuperTip:="Alternates between setting the advanced buttons visible and hidden.", _
                AlternateLabel:=""), _
            Visible:=True, Enabled:=True)
#End If
                
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''' Group 2
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Set CustomGroup = .NewRibbonGroup( _
            Strings.AddControl(ControlID:="Group2", _
                Label:="Advanced Buttons", _
                ScreenTip:="Buttons Optionally Visible", _
                SuperTip:="These buttons are visible on this ribbon only when requested."), _
            True)
        Set CustomSizeToggle = .NewRibbonToggleMso( _
            Strings.AddControl(ControlID:="SizeToggle", _
                Label:="Prefer Small", _
                ScreenTip:="Toggles Button Size", _
                SuperTip:="Toggles between large and small (preferred size for) buttons for this ribbon group.", _
                AlternateLabel:="Prefer Large"), _
            True, True, rdLarge, DefaultImage, False, True)
        
        Set AdvancedOptions = .NewRibbonDropDown( _
            Strings.AddControl(ControlID:="AdvancedDropdown1", _
                Label:="Options", _
                ScreenTip:="Advanced Button Options", _
                SuperTip:="Display options for Advanced Buttons:" & _
                        vbNewLine & "Label Only;" & _
                        vbNewLine & "Image Only;" & _
                        vbNewLine & "Label & Image"), _
            Visible:=True, Enabled:=False)
        AdvancedOptions.AddItem .NewSelectableItem( _
            Strings.AddControl(ControlID:="Group2Dropdown1_LabelOnly", _
                Label:="Label Only", _
                ScreenTip:="Display Label Only", _
                SuperTip:="Displays just the button label when its size is Regular; else both."))
        AdvancedOptions.AddItem .NewSelectableItem( _
            Strings.AddControl(ControlID:="Group2Dropdown1_ImageOnly", _
                Label:="Image Only", _
                ScreenTip:="Display Image Only", _
                SuperTip:="Displays just the button image when its size is Regular; else both."))
        AdvancedOptions.AddItem .NewSelectableItem( _
            Strings.AddControl(ControlID:="Group2Dropdown1_LabelAndImage", _
                Label:="Label & Image", _
                ScreenTip:="Display Label & Image", _
                SuperTip:="Always displays both the button label and image."))
        Set Discard = .NewRibbonDropDown( _
            Strings.AddControl(ControlID:="AdvancedDropdown2", Label:=""), _
            Visible:=False, Enabled:=False)
        Set Discard = .NewRibbonDropDown( _
            Strings.AddControl(ControlID:="AdvancedDropdown3", Label:=""), _
            Visible:=False, Enabled:=False)
        
        Set AdvancedButton1 = .NewRibbonButtonMso( _
            Strings.AddControl(ControlID:="AppLaunchButton1", _
                Label:="1st Advanced Button", _
                ScreenTip:="First Custom Button", _
                SuperTip:="First example of a custom button."), _
            Visible:=True, _
            Enabled:=True, _
            Size:=rdLarge, _
            ImageMso:="RefreshAll", _
            ShowImage:=True, _
            ShowLabel:=True)
        Set AdvancedButton2 = .NewRibbonButtonMso( _
            Strings.AddControl(ControlID:="AppLaunchButton2", _
                Label:="2nd Advanced Button", _
                ScreenTip:="Seccond Custom Button", _
                SuperTip:="Second example of a custom button."), _
            Visible:=True, _
            Enabled:=True, _
            Size:=rdLarge, _
            ImageMso:="Refresh", _
            ShowImage:=True, _
            ShowLabel:=True)
        Set AdvancedButton3 = .NewRibbonButtonMso( _
            Strings.AddControl(ControlID:="AppLaunchButton3", _
                Label:="3rd Advanced Button", _
                ScreenTip:="Third Custom Button", _
                SuperTip:="Third example of a custom button."), _
            Visible:=True, _
            Enabled:=True, _
            Size:=rdLarge, _
            ImageMso:="MacroPlay", _
            ShowImage:=True, _
            ShowLabel:=True)
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''' Default Settings
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ShowAdvancedToggle.OnToggled False
        StandardOptions.OnActionDropDown "", 2
    
        AdvancedOptions.SelectedItemId = "Group2Dropdown1_LabelAndImage"
        AdvancedOptions.IsEnabled = CustomSizeToggle.IsPressed
    End With
    This.ViewModel.Invalidate
    
    Set Initialize = Me
XT: Exit Function
EH: ReraiseError Err, TypeName(Me) & ".Initialize"
    Resume XT
    Resume Next
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''   DON'T EDIT THESE ROUTINES   '''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IResourceManager_GetCurrentUIString(ByVal Name As String) As String
    On Error GoTo EH
    IResourceManager_GetCurrentUIString = Strings.Item(Name)
XT: Exit Function
EH: ReraiseError Err, TypeName(Me) & ".IResourceManager_GetCurrentUIString"
    IResourceManager_GetCurrentUIString = "Name Unknown"
    Resume XT
End Function

Private Function IResourceManager_LoadImage(ByVal ImageId As String)
    On Error GoTo EH
    Set IResourceManager_LoadImage = Nothing ' Sheet1.Pictures(ImageId)
XT: Exit Function
EH: ReraiseError Err, TypeName(Me) & ".IResourceManager_LoadImage"
    Set IResourceManager_LoadImage = Nothing
    Resume XT
End Function

Friend Property Get ViewModel() As RibbonViewModel
    Set ViewModel = This.ViewModel
End Property

Friend Sub ActivateTab()
    This.ViewModel.ActivateTab This.RibbonTab.ID
End Sub
