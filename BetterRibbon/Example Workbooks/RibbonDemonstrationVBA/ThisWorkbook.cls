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
Private MRibbonModel            As RibbonModel

Private Sub Workbook_Activate()
    On Error GoTo EH
    If MRibbonModel Is Nothing Then Set MRibbonModel = New RibbonModel
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, TypeName(Me) & ".Workbook_Activate", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    If ThisWorkbook.Worksheets("Control Strings").Visible <> xlSheetVeryHidden Then _
        ThisWorkbook.Worksheets("Control Strings").Visible = xlSheetVeryHidden
End Sub

Private Sub Workbook_Open()
    On Error GoTo EH
    If DeskTop(False) = "D:\Users\Pieter\Desktop\" _
    Or ThisWorkbook.Path = DeskTop(True) & "Example Workbooks" _
    Or ThisWorkbook.Path = DeskTop(False) & "Example Workbooks" Then _
        MsgBox "Pause for Ctrl-Break to ease debugging." & vbNewLine & vbNewLine & _
               "This message can be disabled by moving the workbook" & vbNewLine & _
               "out of the Desktop folder 'Example Workbooks'.", _
               vbOKOnly, ThisWorkbook.Name
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, TypeName(Me) & ".Workbook_Open", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub
