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
Private MBetterRibbon       As PGSolutions_RibbonDispatcher.IModelFactory
Private CustomGroup         As PGSolutions_RibbonDispatcher.IGroupModel

Private WithEvents Button1  As PGSolutions_RibbonDispatcher.ButtonModel
Attribute Button1.VB_VarHelpID = -1

Private Sub Button1_Clicked(ByVal control As IRibbonControl)
    On Error GoTo EH
    LinksLexerTests.TestAll
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, TypeName(Me) & ".Button2_Clicked", vbOKOnly Or vbInformation
    Resume XT
End Sub

Friend Sub Activate()
    If Not CustomGroup Is Nothing Then CustomGroup.Attach "CustomizableGroup"
    If Not CustomGroup Is Nothing Then CustomGroup.SetShowInactive False
    If Not Button1 Is Nothing Then Button1.Attach "CustomizableButton1"
    If Not Button1 Is Nothing Then Button1.Invalidate
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
                ScreenTip:="", superTip:="", keyTip:=""))
        
        Set Strings = .NewControlStrings2(Label:="Run All Scan/Parse Tests", _
                ScreenTip:="Unit Test for Linka Parser", _
                superTip:="Performs a connection test and then ten" & vbNewLine & _
                          "specific scanning and parsing tests on " & vbNewLine & _
                          "various formulas.", keyTip:="", Description:="")
                
        Set Button1 = .NewButtonModelMso(Strings, "MacroPlay")
    End With
End Sub
