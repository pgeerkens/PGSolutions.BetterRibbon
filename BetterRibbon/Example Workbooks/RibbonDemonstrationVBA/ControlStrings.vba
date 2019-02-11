Attribute VB_Name = "ControlStrings"
'''=======================================================================================
'''                Copyright (c) 2017-2019 Pieter Geerkens
'''
'''     Licensed under the MIT Licence at:
'''             https://github.com/pgeerkens/PGSolutions.BetterRibbon/blob/dev/LICENSE
'''=======================================================================================
Option Explicit
Option Private Module
Private Const ModuleName    As String = "ControlStrings"

Public Property Get Toggle1Strings() As PGSolutions_RibbonDispatcher.IRibbonControlStrings
    Set Toggle1Strings = RibbonDispatcher.NewControlStrings(Label:="Show Inactive", _
            ScreenTip:="Toggles Display of Inactive", _
            SuperTip:="Toggles on/off the display of customizable ribbon controls" & _
                      " that are currently unattached to a Data Source and/or" & _
                      " event sink.")
End Property

Public Property Get Toggle2Strings() As PGSolutions_RibbonDispatcher.IRibbonControlStrings
    Set Toggle2Strings = RibbonDispatcher.NewControlStrings(Label:="Prefer Small Toggles", _
            ScreenTip:="Toggles Large or Small Toggles", _
            SuperTip:="Toggles the size of the activated Toggle" & _
                      " Buttons between Large and Regular.")
End Property

Public Property Get Toggle3Strings() As PGSolutions_RibbonDispatcher.IRibbonControlStrings
    Set Toggle3Strings = RibbonDispatcher.NewControlStrings(Label:="Prefer Small Buttons", _
            ScreenTip:="Toggles Large or Small Buttons", _
            SuperTip:="Toggles the size of the activated Action" & _
                      " Buttons between Large and Regular.")
End Property

Public Property Get Dropdown1Strings() As PGSolutions_RibbonDispatcher.IRibbonControlStrings
    Set Dropdown1Strings = RibbonDispatcher.NewControlStrings(Label:="Image or Label", _
            ScreenTip:="Select Image, Label, or Both", _
            SuperTip:="Controls display of Image, Label, or Both" & _
                      " for the customizable action buttons.")
End Property

Public Property Get Dropdown2Strings() As PGSolutions_RibbonDispatcher.IRibbonControlStrings
    Dropdown2Strings = RibbonDispatcher.NewControlStrings(Label:="Dropdown2Strings", _
            ScreenTip:="Dropdown2Strings ScreenTip", _
            SuperTip:="Dropdown2Strings SuperTip")
End Property

Public Property Get Dropdown3Strings() As PGSolutions_RibbonDispatcher.IRibbonControlStrings
    Dropdown3Strings = RibbonDispatcher.NewControlStrings(Label:="Dropdown3Strings", _
            ScreenTip:="Dropdown3Strings ScreenTip", _
            SuperTip:="Dropdown3Strings SuperTip")
End Property

Public Property Get Button1Strings() As PGSolutions_RibbonDispatcher.IRibbonControlStrings
    Set Button1Strings = RibbonDispatcher.NewControlStrings(Label:="This is cool!", _
            ScreenTip:="Button1 Screentip", _
            SuperTip:="Lots of good things" & vbNewLine & _
                      "can be done here to" & vbNewLine & _
                      "show off a bit.", KeyTip:="", AlternateLabel:="", Description:="")
End Property

Public Property Get Button2Strings() As PGSolutions_RibbonDispatcher.IRibbonControlStrings
    Set Button2Strings = RibbonDispatcher.NewControlStrings(Label:="This is cooler!", _
            ScreenTip:="Button2 Screentip", _
            SuperTip:="Lots of good things" & vbNewLine & _
                      "can be done from hither" & vbNewLine & _
                      " " & vbNewLine & _
                      " " & vbNewLine & _
                      " " & vbNewLine & _
                      " " & vbNewLine & _
                      " " & vbNewLine & _
                      " " & vbNewLine & _
                      "... all the way to yon." & vbNewLine & _
                      "to show off a bit more.")
End Property

Public Property Get Button3Strings() As PGSolutions_RibbonDispatcher.IRibbonControlStrings
    Set Button3Strings = RibbonDispatcher.NewControlStrings(Label:="This is coolest!", _
            ScreenTip:="Button3 Screentip", _
            SuperTip:="What's the weather like where you are?.")
End Property
