VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'''=======================================================================================
'''                Copyright (c) 2017-2019 Pieter Geerkens
'''
'''     Licensed under the MIT Licence at:
'''             https://github.com/pgeerkens/PGSolutions.BetterRibbon/blob/dev/LICENSE
'''=======================================================================================
Option Explicit
Private MRibbonModel    As SingleButton.RibbonModel

Private Sub Workbook_Activate()
    If MRibbonModel Is Nothing Then Set MRibbonModel = New RibbonModel
    MRibbonModel.Activate
End Sub

Private Sub Workbook_Open()
    If DeskTop(False) = "D:\Users\Pieter\Desktop" Then _
        MsgBox "Pause for Ctrl-Break to ease debugging." & vbNewLine & vbNewLine & _
               "This message should only display if you are me.", _
               vbOKOnly, ThisWorkbook.Name
End Sub
