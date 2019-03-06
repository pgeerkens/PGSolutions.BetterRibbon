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
 
''' <summary>Local implementation of IResourceLoader serving control display strings to the dispatcher.</summary
''' <remarks>
''' This implementation performs a lookup into a named data range "ControlStringData" on
''' a _very hidden_ worksheet to find the control strings for custom ribbon controls.
''' </remarks>
Implements IResourceLoader
Private Const DataRangeName As String = "ControlStringData"

''' <summary>Serving Button, ToggleButton, CheckBox, Menu, and Gallery controls. </summary>
Private Function IResourceLoader_GetControlStrings2(ByVal Search As String) As IControlStrings2
    On Error GoTo EH
    Dim RowNo As Double: RowNo = Application.Match(Search, Range(DataRangeName).Columns.Item(1), 0)
    With New RibbonControlStrings
        Set IResourceLoader_GetControlStrings2 = .Initialize2(Label:=Index(RowNo, 3), _
            ScreenTip:=Index(RowNo, 4), SuperTip:=Index(RowNo, 5), KeyTip:=Index(RowNo, 2), _
            Description:=Index(RowNo, 6))
    End With
XT: Exit Function
EH: ErrorUtils.ReRaiseError Err, TypeName(Me) & ".IResourceLoader_GetControlStrings2"
    Resume          ' for debugging only
End Function

''' <summary>Serving all other controls. </summary>
Private Function IResourceLoader_GetControlStrings(ByVal Search As String) As IControlStrings
    On Error GoTo EH
    Dim RowNo As Double: RowNo = Application.Match(Search, Range(DataRangeName).Columns.Item(1), 0)
    With New RibbonControlStrings
        Set IResourceLoader_GetControlStrings = .Initialize(Label:=Index(RowNo, 3), _
            ScreenTip:=Index(RowNo, 4), KeyTip:=Index(RowNo, 2), SuperTip:=Index(RowNo, 5))
    End With
XT: Exit Function
EH: ErrorUtils.ReRaiseError Err, TypeName(Me) & ".IResourceLoader_GetControlStrings"
    Resume          ' for debugging only
End Function

Private Function Index(ByVal RowNo As Double, ByVal ColNo As Double) As Variant
    Index = Application.WorksheetFunction.Index(Range(DataRangeName), RowNo, ColNo)
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
    ErrorUtils.ReRaiseError Err, TypeName(Me) & ".FindStr"
    Resume          ' for debugging only
End Function
