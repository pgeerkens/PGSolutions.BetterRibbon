VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private MRibbonModel            As RibbonModel

Private Sub Workbook_Activate()
    On Error GoTo EH
    If MRibbonModel Is Nothing Then Set MRibbonModel = New RibbonModel
    MRibbonModel.Activate
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, TypeName(Me) & ".Workbook_Activate", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub

Private Sub Workbook_Open()
    On Error GoTo EH
    MsgBox "Pause for Ctrl-Break", vbOKOnly, TypeName(Me)
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, TypeName(Me) & ".Workbook_Open", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub
