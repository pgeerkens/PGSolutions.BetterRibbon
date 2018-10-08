Attribute VB_Name = "Macros"
Option Explicit

Private Const ModuleName As String = "Macros."""

''' <summary>Refeshes ribbon for the active workbook, provided its IRibbonUI is still cached.</summary>
Public Sub ReInitializeRibbon()
    On Error GoTo EH
    RibbonLoader.InitializeRibbon Nothing, ActiveWorkbook
XT: Exit Sub
EH: DisplayError Err, ModuleName & "ReInitializeRibbon"
    Resume XT
    Resume      ' for debugging only
End Sub
