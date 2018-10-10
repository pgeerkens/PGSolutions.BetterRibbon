Attribute VB_Name = "Macros"
Option Explicit

Private Const ModuleName As String = "Macros."

''' <summary>Refeshes ribbon for the this workbook.</summary>
Public Sub ReInitializeRibbon()
    On Error GoTo EH
    AddInHandle.ReInitializeRibbon
XT: Exit Sub
EH: DisplayError Err, ModuleName & "ReInitializeRibbon"
    Resume XT
    Resume      ' for debugging only
End Sub
