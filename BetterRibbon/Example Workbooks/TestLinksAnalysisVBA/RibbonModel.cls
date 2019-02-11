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

Private MRibbonDispatcher   As PGSolutions_RibbonDispatcher.IRibbonDispatcher

Private Button1Strings      As PGSolutions_RibbonDispatcher.IRibbonControlStrings
Private WithEvents Button1  As PGSolutions_RibbonDispatcher.RibbonButtonModel
Attribute Button1.VB_VarHelpID = -1

Private Sub Button1_Clicked(ByVal sender As Variant)
    PublicMacros.TestAll
End Sub

Friend Sub Activate()
    Button1.Attach "CustomizableButton1", Button1Strings
    Button1.Invalidate
End Sub

Private Function RibbonDispatcher() As PGSolutions_RibbonDispatcher.IRibbonDispatcher
    If MRibbonDispatcher Is Nothing Then
        Set MRibbonDispatcher = Application.COMAddIns("PGSolutions.BetterRibbon").Object.NewBetterRibbon()
    End If
    Set RibbonDispatcher = MRibbonDispatcher
End Function

Private Sub Class_Initialize()
    With RibbonDispatcher
        Set Button1Strings = .NewControlStrings(Label:="Run All Scan/Parse Tests", _
                ScreenTip:="Unit Test for Linka Parser", _
                superTip:="Performs a connection test and then ten" & vbNewLine & _
                          "specific scanning and parsing tests on " & vbNewLine & _
                          "various formulas.", keyTip:="", _
                alternateLabel:="", Description:="")
                
        Set Button1 = .NewRibbonButtonModel:  Button1.SetImageMso "MacroPlay": Button1.Invalidate
    End With
End Sub
