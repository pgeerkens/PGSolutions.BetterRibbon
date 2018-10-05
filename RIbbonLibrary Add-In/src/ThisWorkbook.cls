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

Private RibbonUICollection As Collection

Public Function GetRibbonUI(ByVal WkBk As Excel.Workbook) As IRibbonUI
    On Error GoTo EH
    Set GetRibbonUI = RibbonUICollection.Item(WkBk.Path)
XT: Exit Function
EH: ReraiseError Err, TypeName(Me) & ".GetRibbonUI"
    Resume          ' for debugging only
End Function

Friend Function SetRibbonUI(ByVal RibbonUI As IRibbonUI, ByVal WkBk As Excel.Workbook) As IRibbonUI
    On Error GoTo EH
    RibbonUICollection.Add RibbonUI, WkBk.Path
    Set SetRibbonUI = RibbonUI
XT: Exit Function
EH: ReraiseError Err, TypeName(Me) & ".SetRibbonUI"
    Resume          ' for debugging only
End Function

Friend Sub Workbook_Open()
    On Error GoTo EH
    Set RibbonUICollection = New Collection
XT: Exit Sub
EH: DisplayError Err, TypeName(Me) & ".Workbook_Open"
    Resume          ' for debugging only
End Sub
