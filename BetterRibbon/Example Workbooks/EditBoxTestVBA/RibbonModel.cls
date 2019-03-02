VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RibbonModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''=======================================================================================
'''                Copyright (c) 2017-2019 Pieter Geerkens
'''
'''     Licensed under the MIT Licence at:
'''             https://github.com/pgeerkens/PGSolutions.BetterRibbon/blob/dev/LICENSE
'''=======================================================================================
Option Explicit

Private CustomGroup             As PGSolutions_RibbonDispatcher.IGroupModel
Private LabelTest               As PGSolutions_RibbonDispatcher.LabelControlModel
Private WithEvents SplitButton  As PGSolutions_RibbonDispatcher.SplitPressButtonModel
Attribute SplitButton.VB_VarHelpID = -1
Private WithEvents SplitToggle  As PGSolutions_RibbonDispatcher.SplitToggleButtonModel
Attribute SplitToggle.VB_VarHelpID = -1
Private WithEvents Dialog       As PGSolutions_RibbonDispatcher.ButtonModel
Attribute Dialog.VB_VarHelpID = -1
Private WithEvents Button       As PGSolutions_RibbonDispatcher.ButtonModel
Attribute Button.VB_VarHelpID = -1
Private WithEvents EditBox      As PGSolutions_RibbonDispatcher.EditBoxModel
Attribute EditBox.VB_VarHelpID = -1
Private Split2Menu              As PGSolutions_RibbonDispatcher.MenuModel
Private WithEvents Gallery1     As PGSolutions_RibbonDispatcher.StaticGalleryModel
Attribute Gallery1.VB_VarHelpID = -1

Private WithEvents InnerToggle  As PGSolutions_RibbonDispatcher.ToggleModel
Attribute InnerToggle.VB_VarHelpID = -1

Private Sub Button_Clicked(ByVal control As Office.IRibbonControl)
    On Error GoTo EH
    If Not CustomGroup Is Nothing Then CustomGroup.Detach
    Class_Initialize
    Activate
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, TypeName(Me) & ".Button3_Clicked", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub

Private Sub Dialog_Clicked(ByVal control As Office.IRibbonControl)
    MsgBox "This mimics a pop-up dialog for testing", vbOKOnly Or vbQuestion, _
            " Dialog Launch Test"
End Sub

Private Sub EditBox_Edited(ByVal control As IRibbonControl, ByVal Text As String)
    ButtonProcessing.EditBox_Processing Text
End Sub

Private Sub Gallery1_SelectionMade(ByVal control As Office.IRibbonControl, ByVal selectedId As String, ByVal SelectedIndex As Long)
    MsgBox "The seleccted item was: " & selectedId, vbOKOnly, ThisWorkbook.Name
End Sub

Private Sub InnerToggle_Toggled(ByVal control As Office.IRibbonControl, ByVal IsPressed As Boolean)
    MsgBox "From InnerToggle", vbOKOnly
End Sub

Private Sub SplitButton_Clicked(ByVal control As Office.IRibbonControl)
    MsgBox "Pressed!", vbOKOnly, "SplitButton_Clicked"
End Sub

Private Sub SplitToggle_Toggled(ByVal control As Office.IRibbonControl, ByVal IsPressed As Boolean)
    MsgBox "Toggled now: " & SplitToggle.IsPressed, vbOKOnly, "SplitToggle_Toggled"
End Sub

Friend Sub Activate()
    If Not CustomGroup Is Nothing Then CustomGroup.Attach "CustomizableGroup"
    If Not CustomGroup Is Nothing Then CustomGroup.SetShowInactive False
    If Not EditBox Is Nothing Then EditBox.Attach "CustomVbaEditBox1"
    If Not EditBox Is Nothing Then EditBox.Invalidate
    If Not LabelTest Is Nothing Then LabelTest.Attach "CustomVbaLabel1"
    If Not Button Is Nothing Then Button.Attach "CustomizableButton3"
    If Not Dialog Is Nothing Then Dialog.Attach "CustomVbaDialog"
    
    If Not SplitButton Is Nothing Then SplitButton.Attach "CustomVbaSplit1"
    If Not SplitToggle Is Nothing Then SplitToggle.Attach "CustomVbaSplit2"
    
    If Not Gallery1 Is Nothing Then Gallery1.Attach "StaticGallery1"
    If Not Split2Menu Is Nothing Then Split2Menu.Attach "CustomVbaSplit2Menu"
End Sub

Private Sub Class_Initialize()
    On Error GoTo EH
    Dim Strings As PGSolutions_RibbonDispatcher.IControlStrings
    With BetterRibbon
        Set CustomGroup = .NewGroupModel("CustomizableGroup")
        Set LabelTest = .NewLabelControlModel("CustomVbaLabel1")
        
        Set EditBox = .NewEditBoxModel("CustomVbaEditBox1")
        EditBox.Text = "Default"
        
        Set Button = .NewButtonModel("CustomizableButton3") _
                     .SetImage(.NewImageObjectMso("RefreshAll"))
        Set Dialog = .NewButtonModel("CustomVbaDialog")
        Set SplitButton = .NewSplitPressButtonModel("CustomVbaSplit1", _
                        "CustomVbaSplit1Menu", "CustomVbaSplit1Button")
        Set SplitToggle = .NewSplitToggleButtonModel("CustomVbaSplit2", _
                        "CustomVbaSplit2Menu", "CustomVbaSplit2Toggle")
        
        Set Gallery1 = .NewStaticGalleryModel("StaticGallery1")
        Set Split2Menu = .NewMenuModel("CustomVbaSplit2Menu")
        
        Set InnerToggle = SplitToggle.ToggleModel
    End With
XT: Exit Sub
EH: MsgBox Err.Description, vbOKOnly
    Resume Next
End Sub
