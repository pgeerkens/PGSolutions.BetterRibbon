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

Private WithEvents DynamicMenu    As PGSolutions_RibbonDispatcher.DynamicMenuModel
Attribute DynamicMenu.VB_VarHelpID = -1
Private WithEvents DynamicButton  As PGSolutions_RibbonDispatcher.ButtonModel
Attribute DynamicButton.VB_VarHelpID = -1
Private WithEvents DynamicToggle  As PGSolutions_RibbonDispatcher.ToggleModel
Attribute DynamicToggle.VB_VarHelpID = -1
Private WithEvents DynamicGallery As PGSolutions_RibbonDispatcher.StaticGalleryModel
Attribute DynamicGallery.VB_VarHelpID = -1

Private Sub Button_Clicked(ByVal control As Office.IRibbonControl)
    On Error GoTo EH
    If Not CustomGroup Is Nothing Then CustomGroup.Detach
    Class_Initialize
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, TypeName(Me) & ".Button3_Clicked", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub

Private Sub Dialog_Clicked(ByVal control As Office.IRibbonControl)
    MsgBox "This mimics a pop-up dialog for testing", vbOKOnly Or vbQuestion, "Dialog Pop-Up"
End Sub

Private Sub DynamicButton_Clicked(ByVal control As Office.IRibbonControl)
    MsgBox "Pressed!", vbOKOnly, "DynamicButton_Clicked"
End Sub

Private Sub DynamicGallery_SelectionMade(ByVal control As Office.IRibbonControl, ByVal SelectedId As String, ByVal SelectedIndex As Long)
    If Not DynamicGallery Is Nothing Then _
        DynamicGallery.Label = SelectedId & " (#" & SelectedIndex & ")" & ""
    If Not DynamicMenu Is Nothing Then _
        DynamicMenu.Label = SelectedId & " (#" & SelectedIndex & ")" & ""
End Sub

Private Sub DynamicMenu_GetContent(ByVal control As Office.IRibbonControl, ByRef Content As String)
    On Error GoTo EH
    Dim ws As Worksheet: Set ws = Sheets.Item("XmlContent")
    Dim r As Range: Set r = ws.Range("A1")
    Content = Sheets.Item("XmlContent").Range("A1").Value2
XT: Exit Sub
EH: MsgBox Err.Description, vbOKOnly
    Resume Next
End Sub

Private Sub DynamicMenu_ContentLoaded(ByVal control As Office.IRibbonControl)
    On Error GoTo EH
    With ThisWorkbook.ModelServer
        Set DynamicButton = .GetButtonModel("Custom2VbaButton1")
        Set DynamicToggle = .GetToggleModel("Custom2VbaToggleButton2")
        Set DynamicGallery = .GetStaticGalleryModel("Custom2VbaStaticGallery2") _
                             .SetImage(.NewImageObjectMso("FunctionsDateTimeInsertGallery"))
    End With
    DynamicGallery_SelectionMade Nothing, DynamicGallery.SelectedId, DynamicGallery.SelectedIndex
    DynamicGallery.Invalidate
XT: Exit Sub
EH: MsgBox Err.Description, vbOKOnly
    Resume Next
End Sub

Private Sub DynamicToggle_Toggled(ByVal control As Office.IRibbonControl, ByVal IsPressed As Boolean)
    ' content undefined
End Sub

Private Sub EditBox_Edited(ByVal control As IRibbonControl, ByVal Text As String)
    ButtonProcessing.EditBox_Processing Text
End Sub

Private Sub Gallery1_SelectionMade(ByVal control As Office.IRibbonControl, ByVal SelectedId As String, ByVal SelectedIndex As Long)
    MsgBox "The seleccted item was: " & SelectedId, vbOKOnly, ThisWorkbook.Name
End Sub

Private Sub SplitButton_Clicked(ByVal control As Office.IRibbonControl)
    MsgBox "Pressed!", vbOKOnly, "SplitButton_Clicked"
End Sub

Private Sub SplitToggle_Toggled(ByVal control As Office.IRibbonControl, ByVal IsPressed As Boolean)
    MsgBox "Toggled now: " & SplitToggle.IsPressed, vbOKOnly, "SplitToggle_Toggled"
End Sub

Private Sub Class_Initialize()
    On Error GoTo EH
    With ThisWorkbook.ModelServer
        Set CustomGroup = .GetGroupModel("CustomizableGroup")
        Set LabelTest = .GetLabelControlModel("CustomVbaLabel1")
        
        Set EditBox = .GetEditBoxModel("CustomVbaEditBox1")
        EditBox.Text = "Default"
        
        Set Button = .GetButtonModel("CustomizableButton3") _
                     .SetImage(.NewImageObjectMso("RefreshAll"))
        Set Dialog = .GetButtonModel("CustomVbaDialog")
        Set SplitButton = .GetSplitPressButtonModel("CustomVbaSplit1", _
                        "CustomVbaSplit1Menu", "CustomVbaSplit1Button")
        Set SplitToggle = .GetSplitToggleButtonModel("CustomVbaSplit2", _
                        "CustomVbaSplit2Menu", "CustomVbaSplit2Toggle")
        SplitToggle.IsPressed = False
        
        Set Gallery1 = .GetStaticGalleryModel("StaticGallery1")
        Set Split2Menu = .GetMenuModel("CustomVbaSplit2Menu")
        
        Set DynamicMenu = .GetDynamicMenuModel("VbaDynamicMenu1")
    End With
    CustomGroup.Invalidate
XT: Exit Sub
EH: MsgBox Err.Description, vbOKOnly
    Resume Next
End Sub
