Attribute VB_Name = "ButtonProcessing"
'''=======================================================================================
'''                Copyright (c) 2017-2019 Pieter Geerkens
'''
'''     Licensed under the MIT Licence at:
'''             https://github.com/pgeerkens/PGSolutions.BetterRibbon/blob/dev/LICENSE
'''=======================================================================================
Option Explicit
Option Private Module
Private Const COMAddInName  As String = "PGSolutions.BetterRibbon"
Private MBetterRibbon       As PGSolutions_RibbonDispatcher.IModelFactory

Public Function BetterRibbon() As PGSolutions_RibbonDispatcher.IModelFactory
    If MBetterRibbon Is Nothing Then
        Set MBetterRibbon = Application.COMAddIns(COMAddInName).Object _
                .NewBetterRibbon(New ResourceLoader)
    End If
    Set BetterRibbon = MBetterRibbon
End Function

Public Sub EditBox_Processing(ByVal Text As String)
    MsgBox "VBA EditBox edited to value: '" & Text & "'.", vbOKOnly Or vbInformation, ActiveWorkbook.Name
End Sub
