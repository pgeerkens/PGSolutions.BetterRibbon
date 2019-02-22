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
Private Const COMAddInName  As String = "PGSolutions.BetterRibbon"
Private MBetterRibbon       As PGSolutions_RibbonDispatcher.IRibbonDispatcher
Private CustomGroup         As PGSolutions_RibbonDispatcher.GroupModel

Private WithEvents EditBox  As PGSolutions_RibbonDispatcher.EditBoxModel
Attribute EditBox.VB_VarHelpID = -1

Private Sub EditBox_Edited(ByVal sender As Variant, ByVal text As String)
    ButtonProcessing.EditBox_Processing
End Sub

Friend Sub Activate()
    If Not CustomGroup Is Nothing Then CustomGroup.Attach "CustomizableGroup"
    If Not CustomGroup Is Nothing Then CustomGroup.SetShowInactive False
    If Not EditBox Is Nothing Then EditBox.Attach "CustomVbaEditBox1"
    If Not EditBox Is Nothing Then EditBox.Invalidate
End Sub

Private Function BetterRibbon() As PGSolutions_RibbonDispatcher.IRibbonDispatcher
    If MBetterRibbon Is Nothing Then
        Set MBetterRibbon = Application.COMAddIns(COMAddInName).Object.NewBetterRibbon()
    End If
    Set BetterRibbon = MBetterRibbon
End Function

Private Sub Class_Initialize()
    Dim Strings As PGSolutions_RibbonDispatcher.IRibbonControlStrings
    With BetterRibbon
        Set CustomGroup = .NewGroupModel(.NewControlStrings(Label:="VBA Custom Controls"))
        
        Set Strings = .NewControlStrings(Label:="This is cool!", _
                ScreenTip:="VBA-Customized Button Screentip", _
                SuperTip:="This button is completely" & vbNewLine & _
                          "customized within the VBA" & vbNewLine & _
                          "workbook.", keyTip:="", _
                AlternateLabel:="", Description:="")
        
        Set EditBox = .NewEditBoxModel(Strings)
    End With
End Sub
