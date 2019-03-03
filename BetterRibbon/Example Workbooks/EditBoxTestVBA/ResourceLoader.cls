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
    Select Case ControlId
        Case "CustomizableButton3":
            Set IResourceLoader_GetControlStrings2 = BetterRibbon.NewControlStrings2( _
                Label:="Refresh Ribbon", _
                ScreenTip:="Unit Test for Linka Parser", _
                superTip:="Performs Initialize & Attach on the whole Ribbon Tab.", _
                KeyTip:="", Description:="")
        Case "CustomVbaSplit1Menu":
            Set IResourceLoader_GetControlStrings2 = BetterRibbon.NewControlStrings2( _
                Label:="Split1 Menu", _
                ScreenTip:="", _
                superTip:="", _
                KeyTip:="", Description:="")
        Case "CustomVbaSplit1Button":
            Set IResourceLoader_GetControlStrings2 = BetterRibbon.NewControlStrings2( _
                Label:="Split1 Button", _
                ScreenTip:="", _
                superTip:="", _
                KeyTip:="", Description:="")
        Case "CustomVbaSplit2Menu":
            Set IResourceLoader_GetControlStrings2 = BetterRibbon.NewControlStrings2( _
                Label:="Split2 Menu", _
                ScreenTip:="", _
                superTip:="", _
                KeyTip:="", Description:="")
        Case "CustomVbaSplit2Toggle":
            Set IResourceLoader_GetControlStrings2 = BetterRibbon.NewControlStrings2( _
                Label:="Split2 Toggle", _
                ScreenTip:="", _
                superTip:="", _
                KeyTip:="", Description:="")
        Case "CustomVbaDialog":
            Set IResourceLoader_GetControlStrings2 = BetterRibbon.NewControlStrings2( _
                Label:="A Dialog", _
                ScreenTip:="", _
                superTip:="", _
                KeyTip:="", Description:="")
        Case Else:
            Set IResourceLoader_GetControlStrings2 = Nothing
    End Select
End Function

' All other controls
Private Function IResourceLoader_GetControlStrings(ByVal ControlId As String) As PGSolutions_RibbonDispatcher.IControlStrings
    Select Case ControlId
        Case "CustomizableGroup":
            Set IResourceLoader_GetControlStrings = BetterRibbon.NewControlStrings( _
                Label:="VBA Custom Controls", _
                ScreenTip:="", superTip:="", KeyTip:="")
        Case "CustomVbaEditBox1":
            Set IResourceLoader_GetControlStrings = BetterRibbon.NewControlStrings( _
                Label:="An EditBox", _
                ScreenTip:="", _
                superTip:="", _
                KeyTip:="")
        Case "CustomVbaLabel1":
            Set IResourceLoader_GetControlStrings = BetterRibbon.NewControlStrings( _
                Label:="A Label", _
                ScreenTip:="", _
                superTip:="", _
                KeyTip:="")
        Case "CustomVbaSplit1":
            Set IResourceLoader_GetControlStrings = BetterRibbon.NewControlStrings( _
                Label:="SplitPress", _
                ScreenTip:="", _
                superTip:="", _
                KeyTip:="")
        Case "CustomVbaSplit2":
            Set IResourceLoader_GetControlStrings = BetterRibbon.NewControlStrings( _
                Label:="SplitToggle", _
                ScreenTip:="", _
                superTip:="", _
                KeyTip:="")
        Case Else:
            Set IResourceLoader_GetControlStrings = Nothing
    End Select
End Function

Private Function IResourceLoader_GetImage(ByVal Name As String) As Variant
    IResourceLoader_GetImage = "MacroSecurity"
End Function
