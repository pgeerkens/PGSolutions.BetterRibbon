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
    With New RibbonControlStrings
        Set Toggle1Strings = .Initialize(Label:="Show Inactive", _
                ScreenTip:="Toggles Display of Inactive", _
                SuperTip:="Toggles on/off the display of customizable ribbon controls" & _
                          " that are currently unattached to a Data Source and/or" & _
                          " event sink.", _
                AlternateLabel:="")
    End With
End Property

Public Property Get Toggle2Strings() As PGSolutions_RibbonDispatcher.IRibbonControlStrings
    With New RibbonControlStrings
        Set Toggle2Strings = .Initialize(Label:="Prefer Small Toggles", _
                ScreenTip:="Toggles Large or Small Toggles", _
                SuperTip:="Toggles the size of the activated Toggle" & _
                          " Buttons between Large and Regular.")
    End With
End Property

Public Property Get Toggle3Strings() As PGSolutions_RibbonDispatcher.IRibbonControlStrings
    With New RibbonControlStrings
        Set Toggle3Strings = .Initialize(Label:="Prefer Small Buttons", _
                ScreenTip:="Toggles Large or Small Buttons", _
                SuperTip:="Toggles the size of the activated Action" & _
                          " Buttons between Large and Regular.")
    End With
End Property

Public Property Get Dropdown1Strings() As PGSolutions_RibbonDispatcher.IRibbonControlStrings
    With New RibbonControlStrings
        Set Dropdown1Strings = .Initialize(Label:="Image or Label", _
                ScreenTip:="Select Image, Label, or Both", _
                SuperTip:="Controls display of Image, Label, or Both" & _
                          " for the customizable action buttons." & vbNewLine & vbNewLine & _
                          "Only enabled for Regular size buttons.")
    End With
End Property

Public Property Get Dropdown2Strings() As PGSolutions_RibbonDispatcher.IRibbonControlStrings
    With New RibbonControlStrings
        Dropdown2Strings = .Initialize(Label:="Dropdown2Strings", _
                ScreenTip:="Dropdown2Strings ScreenTip", _
                SuperTip:="Dropdown2Strings SuperTip")
    End With
End Property

Public Property Get Dropdown3Strings() As PGSolutions_RibbonDispatcher.IRibbonControlStrings
    With New RibbonControlStrings
        Dropdown3Strings = .Initialize(Label:="Dropdown3Strings", _
                ScreenTip:="Dropdown3Strings ScreenTip", _
                SuperTip:="Dropdown3Strings SuperTip")
    End With
End Property

Public Property Get Button1Strings() As PGSolutions_RibbonDispatcher.IRibbonControlStrings
    With New RibbonControlStrings
        Set Button1Strings = .Initialize(Label:="This is cool!", _
                ScreenTip:="Button1 Screentip", _
                SuperTip:="Lots of good things" & vbNewLine & _
                          "can be done here to" & vbNewLine & _
                          "show off a bit.", KeyTip:="", AlternateLabel:="", Description:="")
    End With
End Property

Public Property Get Button2Strings() As PGSolutions_RibbonDispatcher.IRibbonControlStrings
    With New RibbonControlStrings
        Set Button2Strings = .Initialize(Label:="This is cooler!", _
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
    End With
End Property

Public Property Get Button3Strings() As PGSolutions_RibbonDispatcher.IRibbonControlStrings
    With New RibbonControlStrings
        Set Button3Strings = .Initialize(Label:="This is coolest!", _
                ScreenTip:="Button3 Screentip", _
                SuperTip:="What's the weather like where you are?.")
    End With
End Property
