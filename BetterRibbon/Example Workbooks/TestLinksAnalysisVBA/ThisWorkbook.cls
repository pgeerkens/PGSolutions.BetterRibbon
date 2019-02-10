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
Private MRibbonModel    As TestLinksAnalysis.RibbonModel

Private Sub Workbook_Activate()
    If MRibbonModel Is Nothing Then Set MRibbonModel = New RibbonModel
    MRibbonModel.Activate
End Sub

Private Sub Workbook_Open()
    MsgBox "Pause for Ctrl-Break", vbOKOnly, TypeName(Me)
End Sub
