VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RibbonModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents Button1  As PGSolutions_RibbonDispatcher.RibbonButton
Attribute Button1.VB_VarHelpID = -1
Private Button1Strings      As PGSolutions_RibbonDispatcher.IRibbonControlStrings

Private Sub Button1_Clicked()
    MsgBox "Button1 clicked.", vbOKOnly Or vbInformation, TypeName(Me)
End Sub

Friend Sub Activate()
    With AddInHandle
        Set Button1 = .AttachButton("CustomizableButton1", Button1Strings)
        Button1.SetImageMso "RefreshAll"
    End With
End Sub

Private Sub Class_Initialize()
    With AddInHandle
        Set Button1Strings = .NewControlStrings(Label:="This is cool!", _
                ScreenTip:="Button1 Screentip", _
                SuperTip:="Lots of good things" & vbNewLine & _
                          "can be done here to" & vbNewLine & _
                          "show off a bit.", keyTip:="", _
                AlternateLabel:="", Description:="")
    End With
End Sub

Private Function AddInHandle() As PGSolutions_RibbonDispatcher.IRibbonDispatcher
    Set AddInHandle = Application.COMAddIns("PGSolutions.BetterRibbon").Object
End Function
