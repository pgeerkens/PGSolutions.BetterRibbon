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
Private Function IResourceLoader_GetControlStrings2(ByVal Search As String) As IControlStrings2
    Const ControlStringsRangeName As String = "ControlStringData"
    
    On Error GoTo EH
    With New RibbonControlStrings
        Set IResourceLoader_GetControlStrings2 = .Initialize2( _
                KeyTip:=FindStr(Search, Range(ControlStringsRangeName), 2), _
                Label:=FindStr(Search, Range(ControlStringsRangeName), 3), _
                ScreenTip:=FindStr(Search, Range(ControlStringsRangeName), 5), _
                SuperTip:=FindStr(Search, Range(ControlStringsRangeName), 6), _
                Description:=FindStr(Search, Range(ControlStringsRangeName), 7))
    End With
XT: Exit Function
EH: ErrorUtils.ReRaiseError Err, TypeName(Me) & ".IResourceLoader_GetControlStrings2"
    Resume          ' for debugging only
End Function

' All other controls
Private Function IResourceLoader_GetControlStrings(ByVal Search As String) As IControlStrings
    Const ControlStringsRangeName As String = "ControlStringData"
    
    On Error GoTo EH
    With New RibbonControlStrings
        Set IResourceLoader_GetControlStrings = .Initialize( _
                KeyTip:=FindStr(Search, Range(ControlStringsRangeName), 2), _
                Label:=FindStr(Search, Range(ControlStringsRangeName), 3), _
                ScreenTip:=FindStr(Search, Range(ControlStringsRangeName), 5), _
                SuperTip:=FindStr(Search, Range(ControlStringsRangeName), 6))
    End With
XT: Exit Function
EH: ErrorUtils.ReRaiseError Err, TypeName(Me) & ".BetterRibbon"
    Resume          ' for debugging only
End Function

Private Function IResourceLoader_GetImage(ByVal Name As String) As Variant
    IResourceLoader_GetImage = "MacroSecurity"
End Function

Private Function FindStr(ByVal Search As String, ByVal Range As Range, ByVal ColNo As Integer) As String
    On Error GoTo EH
    Dim v As Variant: v = Application.WorksheetFunction.VLookup(Search, Range, ColNo, False)
    FindStr = IIf(VarType(v) = vbString, v, "")
XT: Exit Function
EH: If Err.Number = 1004 Then Resume Next
    ErrorUtils.ReRaiseError Err, TypeName(Me) & ".BetterRibbon"
    Resume          ' for debugging only
End Function