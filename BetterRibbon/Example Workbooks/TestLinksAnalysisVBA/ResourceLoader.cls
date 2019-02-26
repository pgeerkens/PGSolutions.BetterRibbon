VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ResourceLoader"
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

Implements IResourceLoader

' Button, ToggleButton, CheckBox, Menu, and Gallery controls
Private Function IResourceLoader_GetControlStrings2(ByVal ControlId As String) As PGSolutions_RibbonDispatcher.IControlStrings2
    Select Case (ControlId)
        Case "CustomizableButton1":
            Set IResourceLoader_GetControlStrings2 = BetterRibbon.NewControlStrings2( _
                Label:="Run All Scan/Parse Tests", _
                ScreenTip:="Unit Test for Linka Parser", _
                superTip:="Performs a connection test and then ten" & vbNewLine & _
                          "specific scanning and parsing tests on " & vbNewLine & _
                          "various formulas.", keyTip:="", Description:="")
        Case Else:
            Set IResourceLoader_GetControlStrings2 = Nothing
    End Select
End Function

' All other controls
Private Function IResourceLoader_GetControlStrings(ByVal ControlId As String) As PGSolutions_RibbonDispatcher.IControlStrings
    Select Case (ControlId)
        Case "CustomizableGroup":
            Set IResourceLoader_GetControlStrings = BetterRibbon.NewControlStrings( _
                Label:="VBA Custom Controls", _
                ScreenTip:="", superTip:="", keyTip:="")
        Case Else:
            Set IResourceLoader_GetControlStrings = Nothing
    End Select
End Function

Private Function IResourceLoader_GetImage(ByVal Name As String) As Variant
    IResourceLoader_GetImage = "MacroSecurity"
End Function
