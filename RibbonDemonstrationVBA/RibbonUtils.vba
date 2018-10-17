Attribute VB_Name = "RibbonUtils"
Option Explicit
Option Private Module
Private Const ModuleName As String = "RibbonUtilities."

Public Function AddInHandle() As PGSolutions_RibbonDispatcher.IRibbonDispatcher
    On Error GoTo EH
    Set AddInHandle = Application.COMAddIns("PGSolutions.BetterRibbon").Object
XT: Exit Function
EH: ErrorUtils.ReRaiseError Err, ModuleName & "AddInHandle"
    Resume          ' for debugging only
End Function
