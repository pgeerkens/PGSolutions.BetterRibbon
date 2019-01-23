VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RibbonModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Button1Strings      As PGSolutions_RibbonDispatcher.IRibbonControlStrings
Private Button2Strings      As PGSolutions_RibbonDispatcher.IRibbonControlStrings
Private Button3Strings      As PGSolutions_RibbonDispatcher.IRibbonControlStrings
Private ToggleStrings       As PGSolutions_RibbonDispatcher.IRibbonControlStrings

Private WithEvents Button1  As PGSolutions_RibbonDispatcher.RibbonButton
Attribute Button1.VB_VarHelpID = -1
Private WithEvents Button2  As PGSolutions_RibbonDispatcher.RibbonButton
Attribute Button2.VB_VarHelpID = -1
Private WithEvents Button3  As PGSolutions_RibbonDispatcher.RibbonButton
Attribute Button3.VB_VarHelpID = -1
Private WithEvents Toggle   As RibbonToggleModel
Attribute Toggle.VB_VarHelpID = -1

Private Sub Button1_Clicked(ByVal Sender As Variant)
    On Error GoTo EH
    ButtonProcessing.Button1_Processing ThisWorkbook.Name
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, TypeName(Me) & ".Button1_Clicked", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub

Private Sub Button2_Clicked(ByVal Sender As Variant)
    On Error GoTo EH
    ButtonProcessing.Button2_Processing ThisWorkbook.Name
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, TypeName(Me) & ".Button2_Clicked", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub

Private Sub Button3_Clicked(ByVal Sender As Variant)
    On Error GoTo EH
    ButtonProcessing.Button3_Processing ThisWorkbook.Name
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, TypeName(Me) & ".Button3_Clicked", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub

Private Sub Toggle_Toggled(ByVal Sender As Variant, ByVal IsPressed As Boolean)
    On Error GoTo EH
    RibbonDispatcher.ShowInactive IsPressed
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, TypeName(Me) & ".Toggle_Toggled", vbOKOnly Or vbInformation
    Resume          ' for debugging only
End Sub

Friend Sub Activate()
    On Error GoTo EH
    With RibbonDispatcher
        Set Button1 = .AttachButton("CustomizableButton1", Button1Strings)
        Button1.SetImageMso "RefreshAll"
        Set Button2 = .AttachButton("CustomizableButton2", Button2Strings)
        Button2.SetImageMso "Refresh"
        Set Button3 = .AttachButton("CustomizableButton3", Button3Strings)
        Button3.SetImageMso "MacroPlay"
    End With
    With New RibbonToggleModel
        Set Toggle = .Attach("CustomVbaToggle", ToggleStrings)
    End With
    RibbonDispatcher.InvalidateCustomControlsGroup
XT: Exit Sub
EH: ErrorUtils.ReRaiseError Err, TypeName(Me) & ".Activate"
    Resume          ' for debugging only
End Sub

Private Sub Class_Initialize()
    On Error GoTo EH
    With RibbonDispatcher
        Set Button1Strings = .NewControlStrings(Label:="This is cool!", _
                ScreenTip:="Button1 Screentip", _
                SuperTip:="Lots of good things" & vbNewLine & _
                          "can be done here to" & vbNewLine & _
                          "show off a bit.", KeyTip:="", AlternateLabel:="", Description:="")
        Set Button2Strings = .NewControlStrings(Label:="This is cooler!", _
                ScreenTip:="Button2 Screentip", _
                SuperTip:="Lots of good things" & vbNewLine & _
                          "can be done from hither" & vbNewLine & _
                          " " & vbNewLine & _
                          " " & vbNewLine & _
                          " " & vbNewLine & _
                          " " & vbNewLine & _
                          " " & vbNewLine & _
                          " " & vbNewLine & _
                          "... all the way to yon." & vbNewLine & _
                          "to show off a bit more.")
        Set Button3Strings = .NewControlStrings(Label:="This is coolest!", _
                ScreenTip:="Button3 Screentip", _
                SuperTip:="What's the weather like where you are?.")
        Set ToggleStrings = .NewControlStrings(Label:="Show Inactive", _
                ScreenTip:="Toggles Display of Inactive", _
                SuperTip:="Toggles on/off the display of customizable ribbon controls" & _
                          " that are currently unattached to a Data Source and/or" & _
                          "event sink.")
    End With
XT: Exit Sub
EH: ErrorUtils.ReRaiseError Err, TypeName(Me) & ".Class_Initialize"
    Resume          ' for debugging only
End Sub
