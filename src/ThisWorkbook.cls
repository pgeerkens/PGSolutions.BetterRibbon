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
Private Const ModuleName        As String = "ThisWorkbook."

Private MRibbonModel            As RibbonDemonstration.RibbonModel

Private Sub Workbook_Activate()
    On Error GoTo EH
    MRibbonModel.Activate
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, ModuleName & "Workbook_Activate", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub

Private Sub Workbook_Open()
    On Error GoTo EH
    Set MRibbonModel = New RibbonModel
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, ModuleName & "Workbook_Open", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub
