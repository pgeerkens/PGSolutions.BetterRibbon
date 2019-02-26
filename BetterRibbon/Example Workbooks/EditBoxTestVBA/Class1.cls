VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
        Case "CustomizableButton3":
            Set IResourceLoader_GetControlStrings2 = BetterRibbon.NewControlStrings2( _
                Label:="Refresh Ribbon", _
                ScreenTip:="Refreshes the Ribbon", _
                superTip:="Re-runs the methods Class_Initialize and Activate.", _
                KeyTip:="", Description:="")
        Case "CustomVbaSplit1Button":
            Set IResourceLoader_GetControlStrings2 = BetterRibbon.NewControlStrings2( _
                Label:="Button Label", _
                ScreenTip:="", superTip:="", KeyTip:="", Description:="")
        Case "CustomVbaSplit1Menu":
            Set IResourceLoader_GetControlStrings2 = BetterRibbon.NewControlStrings2( _
                Label:="Menu Label", _
                ScreenTip:="", superTip:="", KeyTip:="", Description:="")
        Case "CustomVbaSplit2Toggle":
            Set IResourceLoader_GetControlStrings2 = BetterRibbon.NewControlStrings2( _
                Label:="Toggle Label", _
                ScreenTip:="", superTip:="", KeyTip:="", Description:="")
        Case "CustomVbaSplit2Menu":
            Set IResourceLoader_GetControlStrings2 = BetterRibbon.NewControlStrings2( _
                Label:="Menu Label", _
                ScreenTip:="", superTip:="", KeyTip:="", Description:="")
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
                ScreenTip:="", superTip:="", KeyTip:="")
        Case "CustomVbaEditBox1":
            Set IResourceLoader_GetControlStrings = BetterRibbon.NewControlStrings( _
                Label:="Enter Text", _
                ScreenTip:="Enter up to 15 characters", _
                superTip:="Accepts, and echoes back, up to 15 characters entered by the user.", _
                KeyTip:="")
        Case "CustomVbaLabel1":
            Set IResourceLoader_GetControlStrings = BetterRibbon.NewControlStrings( _
                Label:="A Label Here", _
                ScreenTip:="Label ScreenTip", _
                superTip:="Some text for the Label SuperTip.", KeyTip:="")
        Case "CustomVbaDialog":
            Set IResourceLoader_GetControlStrings = BetterRibbon.NewControlStrings( _
                Label:="Pop-Up Dialog", _
                ScreenTip:="Used to Edit the Ribbon", _
                superTip:="Pops-up a Ribbon Editor Dialog.", _
                KeyTip:="")
        Case "CustomVbaSplit1":
            Set IResourceLoader_GetControlStrings = BetterRibbon.NewControlStrings( _
                Label:="Split Label", ScreenTip:="", superTip:="", KeyTip:="")
        Case "CustomVbaSplit2":
            Set IResourceLoader_GetControlStrings = BetterRibbon.NewControlStrings( _
                Label:="Split Label", ScreenTip:="", superTip:="", KeyTip:="")
        Case Else:
            Set IResourceLoader_GetControlStrings = Nothing
    End Select
End Function

Private Function IResourceLoader_GetImage(ByVal Name As String) As Variant
    IResourceLoader_GetImage = "MacroSecurity"
End Function
