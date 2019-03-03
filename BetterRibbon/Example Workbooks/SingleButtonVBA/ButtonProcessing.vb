Attribute VB_Name = "ButtonProcessing"
'''=======================================================================================
'''                Copyright (c) 2017-2019 Pieter Geerkens
'''
'''     Licensed under the MIT Licence at:
'''             https://github.com/pgeerkens/PGSolutions.BetterRibbon/blob/dev/LICENSE
'''=======================================================================================
Option Explicit
Option Private Module
Public Const COMAddInName  As String = "PGSolutions.BetterRibbon"
Private MBetterRibbon       As PGSolutions_RibbonDispatcher.IModelFactory

Public Function BetterRibbon() As PGSolutions_RibbonDispatcher.IModelFactory
    If MBetterRibbon Is Nothing Then
        Set MBetterRibbon = Application.COMAddIns(COMAddInName).Object _
                .NewBetterRibbon(New ResourceLoader)
    End If
    Set BetterRibbon = MBetterRibbon
End Function

Public Sub Button1_Processing()
    MsgBox "VBA CustomButton clicked.", vbOKOnly Or vbInformation, ActiveWorkbook.Name
End Sub
