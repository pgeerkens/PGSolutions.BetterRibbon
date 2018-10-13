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
Private Const ModuleName    As String = "ThisWorkbook."

Private WithEvents Button1  As RibbonDispatcherX.RibbonButton
Attribute Button1.VB_VarHelpID = -1
Private WithEvents Button2  As RibbonDispatcherX.RibbonButton
Attribute Button2.VB_VarHelpID = -1
Private WithEvents Button3  As RibbonDispatcherX.RibbonButton
Attribute Button3.VB_VarHelpID = -1

Private Button1Strings      As RibbonDispatcherX.IRibbonTextLanguageControl
Private Button2Strings      As RibbonDispatcherX.IRibbonTextLanguageControl
Private Button3Strings      As RibbonDispatcherX.IRibbonTextLanguageControl

Private Sub Button1_Clicked()
    On Error GoTo EH
    MsgBox "Activation message from Button1!", vbOKOnly Or vbInformation, ModuleName
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, ModuleName & "AppButton1_Clicked", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub

Private Sub Button2_Clicked()
    On Error GoTo EH
    MsgBox "Activation message from Button2!", vbOKOnly Or vbInformation, ModuleName
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, ModuleName & "AppButton1_Clicked", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub

Private Sub Button3_Clicked()
    On Error GoTo EH
    MsgBox "Activation message from Button3!", vbOKOnly Or vbInformation, ModuleName
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, ModuleName & "AppButton1_Clicked", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub

Friend Sub Activate()
    On Error GoTo EH
    With AddInHandle
        Set Button1 = .AttachButton("CustomizableButton1", Button1Strings)
        Button1.SetImageMso "RefreshAll"
        Set Button2 = .AttachButton("CustomizableButton2", Button2Strings)
        Button2.SetImageMso "Refresh"
        Set Button3 = .AttachButton("CustomizableButton3", Button3Strings)
        Button3.SetImageMso "MacroPlay"
    End With
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, ModuleName & "Workbook_Activate", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub

Private Sub Class_Initialize()
    On Error GoTo EH
    With AddInHandle
        Set Button1Strings = .RibbonFactory.NewControlStrings(Label:="This is cool!", _
                 ScreenTip:="Button1 Screentip", _
                 SuperTip:="Lots of good things" & vbNewLine & _
                            "can be done here to" & vbNewLine & _
                            "show off a bit.", KeyTip:="", AlternateLabel:="", Description:="")
        Set Button2Strings = .RibbonFactory.NewControlStrings(Label:="This is cooler!", _
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
        Set Button3Strings = .RibbonFactory.NewControlStrings(Label:="This is coolest!", _
                 ScreenTip:="Button3 Screentip", _
                 SuperTip:="What's the weather like where you are?.")
    End With
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, ModuleName & "Workbook_Activate", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub
