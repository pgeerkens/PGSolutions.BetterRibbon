VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Const ModuleName        As String = "ThisWorkbook."

Private WithEvents Button1      As RibbonDispatcherX.RibbonButton
Attribute Button1.VB_VarHelpID = -1
Private WithEvents Button2      As RibbonDispatcherX.RibbonButton
Attribute Button2.VB_VarHelpID = -1
Private WithEvents Button3      As RibbonDispatcherX.RibbonButton
Attribute Button3.VB_VarHelpID = -1

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

Private Sub Workbook_Activate()
    On Error GoTo EH
    With AddInHandle
        Set Button1 = .AttachProxy("CustomizableButton1", _
            .RibbonFactory.NewLanguageControl("This is cool!", _
                 "Button1 Screentip", _
                 "Lots of good things" & vbNewLine & _
                 "can be done here to" & vbNewLine & _
                 "show off a bit.", KeyTip:="", AlternateLabel:="", Description:=""))
        Button1.SetImageMso "RefreshAll"
        Set Button2 = .AttachProxy("CustomizableButton2", _
            .RibbonFactory.NewLanguageControl("This is cool!", _
                 "Button2 Screentip", _
                 "Lots of good things" & vbNewLine & _
                 "can be done from hither" & vbNewLine & _
                 " " & vbNewLine & _
                 " " & vbNewLine & _
                 " " & vbNewLine & _
                 " " & vbNewLine & _
                 " " & vbNewLine & _
                 " " & vbNewLine & _
                 "... all the way to yon." & vbNewLine & _
                 "show off a bit.", KeyTip:="", AlternateLabel:="", Description:=""))
        Button2.SetImageMso "Refresh"
        Set Button3 = .AttachProxy("CustomizableButton3", _
            .RibbonFactory.NewLanguageControl("This is cool!", _
                 "Button3 Screentip", _
                 "What's the weather like where you are?.", KeyTip:="", AlternateLabel:="", Description:=""))
        Button3.SetImageMso "MacroPlay"
    End With
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, ModuleName & "Workbook_Activate", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub

Private Sub Workbook_Open()
    On Error GoTo EH
            
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, ModuleName & "Workbook_Open", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub
