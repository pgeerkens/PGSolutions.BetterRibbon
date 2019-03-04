Attribute VB_Name = "ButtonProcessing"
'''=======================================================================================
'''                Copyright (c) 2017-2019 Pieter Geerkens
'''
'''     Licensed under the MIT Licence at:
'''             https://github.com/pgeerkens/PGSolutions.BetterRibbon/blob/dev/LICENSE
'''=======================================================================================
Option Explicit
Option Private Module

Public Sub Button1_Processing()
    MsgBox "VBA CustomButton clicked.", vbOKOnly Or vbInformation, ActiveWorkbook.Name
End Sub
