Attribute VB_Name = "RibbonLoader"
Option Explicit

Private Const MModuleName   As String = "RibbonLoader."
Private MRibbonUI           As Office.IRibbonUI

''' <summary>Returns whether or not to display popups during load, for easier debugging.</summary>
Public Const ShowAlerts     As Boolean = True

''' <summary>EventHandler for RibbonLoad, initializing the ViewModel and Model.</summary>
''' <param name="RibbonUI">An IRibbonUI for the ribbon just loaded.</param>
Public Sub OnRibbonLoad(ByVal RibbonUI As Office.IRibbonUI)
    On Error GoTo EH
    If ShowAlerts Then DisplayAlert "OnRibbonLoad"
    
    Set MRibbonUI = SetRibbonUI(RibbonUI, ThisWorkbook)
    InitializeRibbon
    
XT: Exit Sub
EH: DisplayError Err, MModuleName & "OnRibbonLoad"
    Resume XT
    Resume      ' for debugging only
End Sub

Public Sub InitializeRibbon()
    On Error GoTo EH
    If MRibbonUI Is Nothing Then Set MRibbonUI = GetRibbonUI(ThisWorkbook)
    If MRibbonUI Is Nothing Then Err.Raise -1, "", "RibbonUI is Nothing"
    With New RibbonModel
        Set ThisWorkbook.RibbonModel = .Initialize(MRibbonUI)
    End With

XT: Exit Sub
EH: ReraiseError Err, MModuleName & "InitializeRibbon"
    Resume XT
    Resume      ' for debugging only
End Sub
