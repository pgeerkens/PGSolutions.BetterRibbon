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
Private Const COMAddInName      As String = "PGSolutions.BetterRibbon"
Private MBetterRibbon           As PGSolutions_RibbonDispatcher.IModelFactory

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
End Sub

Private Function BetterRibbon() As PGSolutions_RibbonDispatcher.IModelFactory
    If MBetterRibbon Is Nothing Then
        Set MBetterRibbon = Application.COMAddIns(COMAddInName).Object.NewBetterRibbon()
    End If
    Set BetterRibbon = MBetterRibbon
End Function

Private Sub Class_Initialize()
    Dim Strings As PGSolutions_RibbonDispatcher.IControlStrings
    With BetterRibbon
        Set CustomGroup = .NewGroupModel(.NewControlStrings(Label:="VBA Custom Controls", _
                ScreenTip:="", SuperTip:="", KeyTip:=""))
        
        Set LabelTest = .NewLabelModel(.NewControlStrings(Label:="A Label Here", _
                ScreenTip:="Label ScreenTip", _
                SuperTip:="Some text for the Label SuperTip.", KeyTip:=""))
        
        Set EditBox = .NewEditBoxModel(.NewControlStrings(Label:="Enter Text", _
                ScreenTip:="Enter up to 15 characters", _
                SuperTip:="Accepts, and echoes back, up to 15 characters entered by the user.", _
                KeyTip:=""))
        EditBox.Text = "Default"
        
        Set Button = .NewButtonModelMso(.NewControlStrings2(Label:="Refresh Ribbon", _
                ScreenTip:="Refreshes the Ribbon", _
                SuperTip:="Re-runs the methods Class_Initialize and Activate.", _
                KeyTip:="", Description:=""), "RefreshAll")
                
        Set Dialog = .NewButtonModel(.NewControlStrings2(Label:="Pop-Up Dialog", _
                ScreenTip:="Used to Edit the Ribbon", _
                SuperTip:="Pops-up a Ribbon Editor Dialog.", _
                KeyTip:="", Description:=""))
                
        Set SplitButton = .NewSplitPressButtonModel( _
            .NewControlStrings(Label:="Split Label", ScreenTip:="", SuperTip:="", KeyTip:=""), _
            .NewControlStrings2(Label:="Button Label", _
                    ScreenTip:="", SuperTip:="", KeyTip:="", Description:=""), _
            .NewControlStrings2(Label:="Menu Label", _
                     ScreenTip:="", SuperTip:="", KeyTip:="", Description:=""))
                
        Set SplitToggle = .NewSplitToggleButtonModel( _
            .NewControlStrings(Label:="Split Label", ScreenTip:="", SuperTip:="", KeyTip:=""), _
            .NewControlStrings2(Label:="Toggle Label", _
                     ScreenTip:="", SuperTip:="", KeyTip:="", Description:=""), _
            .NewControlStrings2(Label:="Menu Label", _
                     ScreenTip:="", SuperTip:="", KeyTip:="", Description:=""))
            
    End With
End Sub
