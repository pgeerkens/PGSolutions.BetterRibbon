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
Private MRibbonModel    As RibbonModel

Private Sub Workbook_Activate()
    Register
    If MRibbonModel Is Nothing Then Set MRibbonModel = New RibbonModel
End Sub

'Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
'    MsgBox IIf(MRibbonModel Is Nothing, "Now nothing!", "Fine!"), vbOKOnly, "BeforeSave"
'End Sub

'Private Sub Workbook_AfterSave(ByVal Success As Boolean)
'    MsgBox IIf(MRibbonModel Is Nothing, "Now nothing!", "Fine!"), vbOKOnly, "AfterSave"
'End Sub

Private Sub Workbook_Open()
    If DeskTop(False) = "D:\Users\Pieter\Desktop\" _
    Or ThisWorkbook.Path = DeskTop(True) & "Example Workbooks" _
    Or ThisWorkbook.Path = DeskTop(False) & "Example Workbooks" Then _
        MsgBox "Pause for Ctrl-Break to ease debugging." & vbNewLine & vbNewLine & _
               "This message can be disabled by moving the workbook" & vbNewLine & _
               "out of the Desktop folder 'Example Workbooks'.", _
               vbOKOnly, ThisWorkbook.Name
End Sub
