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
Private MModelServer        As PGSolutions_RibbonDispatcher.IModelServer

Public Function ModelServer() As PGSolutions_RibbonDispatcher.IModelServer
    If MModelServer Is Nothing Then
        Set MModelServer = Application.COMAddIns(COMAddInName).Object _
                .NewBetterRibbon(New ResourceLoader)
    End If
    Set ModelServer = MModelServer
End Function

Public Sub Register()
    Application.COMAddIns(COMAddInName).Object.RegisterWorkbook ThisWorkbook.Name
End Sub

Public Sub EditBox_Processing(ByVal Text As String)
    MsgBox "VBA EditBox edited to value: '" & Text & "'.", vbOKOnly Or vbInformation, ActiveWorkbook.Name
End Sub
