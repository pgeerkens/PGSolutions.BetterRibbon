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
Private LabelTest               As PGSolutions_RibbonDispatcher.LabelModel
Private SplitButton  As PGSolutions_RibbonDispatcher.SplitPressButtonModel
Private SplitToggle  As PGSolutions_RibbonDispatcher.SplitToggleButtonModel
Attribute SplitToggle.VB_VarHelpID = -1
Private WithEvents Dialog       As PGSolutions_RibbonDispatcher.ButtonModel
Attribute Dialog.VB_VarHelpID = -1
Private WithEvents Button       As PGSolutions_RibbonDispatcher.ButtonModel
Attribute Button.VB_VarHelpID = -1
Private WithEvents EditBox      As PGSolutions_RibbonDispatcher.EditBoxModel
Attribute EditBox.VB_VarHelpID = -1

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
End Sub

Private Sub Class_Initialize()
    Dim Strings As PGSolutions_RibbonDispatcher.IControlStrings
    With BetterRibbon
        Set CustomGroup = .NewGroupModel("CustomizableGroup")
        Set LabelTest = .NewLabelModel("CustomVbaLabel1")
        
        Set EditBox = .NewEditBoxModel("CustomVbaEditBox1")
        EditBox.Text = "Default"
        
        Set Button = .NewButtonModelMso("CustomizableButton3", "RefreshAll")
        Set Dialog = .NewButtonModel("CustomVbaDialog")
        Set SplitButton = .NewSplitPressButtonModel( _
            "CustomVbaSplit1", "CustomVbaSplit1Menu", "CustomVbaSplit1Button")
        Set SplitToggle = .NewSplitToggleButtonModel( _
            "CustomVbaSplit2", "CustomVbaSplit2Menu", "CustomVbaSplit2Toggle")
            
    End With
End Sub
