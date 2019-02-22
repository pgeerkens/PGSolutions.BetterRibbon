Attribute VB_Name = "ButtonProcessing"
'''=======================================================================================
'''                Copyright (c) 2017-2019 Pieter Geerkens
'''
'''     Licensed under the MIT Licence at:
'''             https://github.com/pgeerkens/PGSolutions.BetterRibbon/blob/dev/LICENSE
'''=======================================================================================
Option Explicit
Option Private Module

Public Sub EditBox_Processing(ByVal Text As String)
    MsgBox "VBA EditBox edited to value: '" & Text & "'.", vbOKOnly Or vbInformation, ActiveWorkbook.Name
End Sub
