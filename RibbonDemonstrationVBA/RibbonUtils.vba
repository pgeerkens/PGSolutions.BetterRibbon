Attribute VB_Name = "RibbonUtils"
Option Explicit
Option Private Module
Private Const ModuleName As String = "RibbonUtilities."

Private RibbonDispatcher As PGSolutions_RibbonDispatcher.IRibbonDispatcher

Public Property Get AddInHandle() As PGSolutions_RibbonDispatcher.IRibbonDispatcher
    If RibbonDispatcher Is Nothing Then Set RibbonDispatcher = NewHandle.NewBetterRibbon()
    Set AddInHandle = RibbonDispatcher
End Property

Private Function NewHandle() As PGSolutions_RibbonDispatcher.IBetterRibbon
    On Error GoTo EH
    Set NewHandle = Application.COMAddIns("PGSolutions.BetterRibbon").Object
XT: Exit Function
EH: ErrorUtils.ReRaiseError Err, ModuleName & "NewHandle"
    Resume          ' for debugging only
End Function
