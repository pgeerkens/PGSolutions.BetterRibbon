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

Private Sub Class_Initialize()
    On Error GoTo EH
    Dim Strings As PGSolutions_RibbonDispatcher.IControlStrings
    With ThisWorkbook.ModelServer
        Set CustomGroup = .GetGroupModel("CustomizableGroup")
        CustomGroup.SetShowInactive False
        Set Button1 = .GetButtonModel("CustomizableButton1") _
                      .SetImage(.NewImageObjectMso("MacroPlay"))
    End With
XT: Exit Sub
EH: MsgBox Err.Description, vbOKOnly Or vbInformation
    Resume Next
End Sub
