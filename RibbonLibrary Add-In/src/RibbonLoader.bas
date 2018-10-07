Attribute VB_Name = "RibbonLoader"
Option Explicit

Private Const MModuleName   As String = "RibbonLoader."

''' <summary>EventHandler for RibbonLoad, initializing the ViewModel and Model.</summary>
''' <param name="RibbonUI">An IRibbonUI for the ribbon just loaded.</param>
Public Sub OnRibbonLoad(ByVal RibbonUI As Office.IRibbonUI)
    On Error GoTo EH
    InitializeRibbon AddInHandle.SetRibbonUI(RibbonUI, ActiveWorkbook.Path), ActiveWorkbook
XT: Exit Sub
EH: DisplayError Err, MModuleName & "OnRibbonLoad"
    Resume XT
    Resume      ' for debugging only
End Sub

''' <summary>Request the supplied IRibbonWorkbook to initialize its RibbonModel.</summary>
''' <param name="RibbonUI">An IRibbonUI for the ribbon just loaded.</param>
''' <param name="WkBk">The IRibbonWorkbook for the ribbon just loaded.</param>
Public Sub InitializeRibbon(ByVal RibbonUI As IRibbonUI, ByVal WkBk As IRibbonWorkbook)
    On Error GoTo EH
    If RibbonUI Is Nothing Then Set RibbonUI = AddInHandle.GetRibbonUI(WkBk.Path)
    If RibbonUI Is Nothing Then Err.Raise 4, WkBk.Path, "RibbonUI is Nothing"
    
    WkBk.InitializeRibbonModel RibbonUI
XT: Exit Sub
EH: ReraiseError Err, MModuleName & "InitializeRibbon"
    Resume      ' for debugging only
End Sub

''' <summary>Refeshes ribbon for the active workbook, provided its IRibbonUI is still cached.</summary>
Public Sub ReInitializeRibbon()
    On Error GoTo EH
    RibbonLoader.InitializeRibbon Nothing, ActiveWorkbook
XT: Exit Sub
EH: DisplayError Err, ModuleName & "ReInitializeRibbon"
    Resume XT
    Resume      ' for debugging only
End Sub
