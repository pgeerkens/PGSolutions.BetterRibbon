Attribute VB_Name = "RibbonLoader"
Option Explicit

Private Const mModuleName   As String = "RibbonLoader."

''' <summary>Returns whether or not to display popups during load, for easier debugging.</summary>
Public Const ShowAlerts     As Boolean = True

''' <summary>EventHandler for RibbonLoad, initializing the ViewModel and Model.</summary>
''' <param name="RibbonUI">An IRibbonUI for the ribbon just loaded.</param>
Public Sub OnRibbonLoad(ByVal RibbonUI As Office.IRibbonUI)
    On Error GoTo EH
    If ShowAlerts Then DisplayAlert "OnRibbonLoad"
    
    With New RibbonModel
        Set ThisWorkbook.RibbonModel = .Initialize(RibbonUI)
        .ActivateTab
    End With
XT: Exit Sub
EH: DisplayError Err, mModuleName & "OnRibbonLoad"
    Resume XT
    Resume      ' for debugging only
End Sub
