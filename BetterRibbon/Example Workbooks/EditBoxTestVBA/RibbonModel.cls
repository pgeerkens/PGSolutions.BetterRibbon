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

Private WithEvents DynamicMenu    As PGSolutions_RibbonDispatcher.DynamicMenuModel
Attribute DynamicMenu.VB_VarHelpID = -1
'Private DynamicMenu     As PGSolutions_RibbonDispatcher.DynamicMenuModel

Private WithEvents DynamicButton  As PGSolutions_RibbonDispatcher.ButtonModel
Attribute DynamicButton.VB_VarHelpID = -1
Private WithEvents DynamicToggle  As PGSolutions_RibbonDispatcher.ToggleModel
Attribute DynamicToggle.VB_VarHelpID = -1
Private WithEvents DynamicGallery As PGSolutions_RibbonDispatcher.StaticGalleryModel
Attribute DynamicGallery.VB_VarHelpID = -1

Private Sub Button_Clicked(ByVal Control As Office.IRibbonControl)
    On Error GoTo EH
    If Not CustomGroup Is Nothing Then CustomGroup.Detach
    Class_Initialize
    Activate
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, TypeName(Me) & ".Button3_Clicked", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub

Private Sub Dialog_Clicked(ByVal Control As Office.IRibbonControl)
    MsgBox "This mimics a pop-up dialog for testing", vbOKOnly Or vbQuestion, _
            " Dialog Launch Test"
End Sub

Private Sub DynamicButton_Clicked(ByVal Control As Office.IRibbonControl)
    MsgBox "Pressed!", vbOKOnly, "DynamicButton_Clicked"
End Sub

Private Sub DynamicGallery_SelectionMade(ByVal Control As Office.IRibbonControl, ByVal SelectedId As String, ByVal SelectedIndex As Long)
    MsgBox "Selected # " & SelectedIndex & ": " & SelectedId, vbOKOnly, "DynamicGallery_SelectionMade"
End Sub

Private Sub DynamicMenu_GetContent(ByVal Control As Office.IRibbonControl, ByRef Content As String)
    On Error GoTo EH
    Dim ws As Worksheet: Set ws = Sheets.Item("XmlContent")
    Dim r As Range: Set r = ws.Range("A1")
    Content = Sheets.Item("XmlContent").Range("A1").Value2
XT: Exit Sub
EH: MsgBox Err.Description, vbOKOnly
    Resume Next
End Sub

Private Sub DynamicMenu_ContentLoaded(ByVal Control As Office.IRibbonControl)
    On Error GoTo EH
    With BetterRibbon
        Set DynamicButton = .NewButtonModel("Custom2VbaButton1")
        Set DynamicToggle = .NewToggleModel("Custom2VbaToggleButton2")
        Set DynamicGallery = .NewStaticGalleryModel("Custom2VbaStaticGallery2")
    End With
    
    If Not DynamicButton Is Nothing Then DynamicButton.Attach "Custom2VbaButton1"
    If Not DynamicToggle Is Nothing Then DynamicToggle.Attach "Custom2VbaToggleButton2"
    If Not DynamicGallery Is Nothing Then DynamicGallery.Attach "Custom2VbaStaticGallery2"
XT: Exit Sub
EH: MsgBox Err.Description, vbOKOnly
    Resume Next
End Sub

Private Sub DynamicToggle_Toggled(ByVal Control As Office.IRibbonControl, ByVal IsPressed As Boolean)
    MsgBox "Toggled to " & IsPressed, vbOKOnly, "DynamicToggle_Toggled"
End Sub

Private Sub EditBox_Edited(ByVal Control As IRibbonControl, ByVal Text As String)
    ButtonProcessing.EditBox_Processing Text
End Sub

Private Sub Gallery1_SelectionMade(ByVal Control As Office.IRibbonControl, ByVal SelectedId As String, ByVal SelectedIndex As Long)
    MsgBox "The seleccted item was: " & SelectedId, vbOKOnly, ThisWorkbook.Name
End Sub

Private Sub InnerToggle_Toggled(ByVal Control As Office.IRibbonControl, ByVal IsPressed As Boolean)
    MsgBox "From InnerToggle", vbOKOnly
End Sub

Private Sub SplitButton_Clicked(ByVal Control As Office.IRibbonControl)
    MsgBox "Pressed!", vbOKOnly, "SplitButton_Clicked"
End Sub

Private Sub SplitToggle_Toggled(ByVal Control As Office.IRibbonControl, ByVal IsPressed As Boolean)
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
    
    If Not DynamicMenu Is Nothing Then DynamicMenu.Attach "VbaDynamicMenu1"
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
        
        Set DynamicMenu = .NewDynamicMenuModel("VbaDynamicMenu1")
    End With
    
    Activate
XT: Exit Sub
EH: MsgBox Err.Description, vbOKOnly
    Resume Next
End Sub
