Attribute VB_Name = "RibbonUtils"
Option Explicit
Option Private Module
Private Const ModuleName    As String = "RibbonUtilities."

Private MRibbonDispatcher   As PGSolutions_RibbonDispatcher.IRibbonDispatcher

Public Property Get RibbonDispatcher() As PGSolutions_RibbonDispatcher.IRibbonDispatcher
    On Error GoTo EH
    If MRibbonDispatcher Is Nothing Then
        Set MRibbonDispatcher = Application.COMAddIns("PGSolutions.BetterRibbon").Object.NewBetterRibbon()
    End If
    Set RibbonDispatcher = MRibbonDispatcher
XT: Exit Property
EH: ErrorUtils.ReRaiseError Err, ModuleName & ".RibbonDispatcher"
    Resume          ' for debugging only
End Property