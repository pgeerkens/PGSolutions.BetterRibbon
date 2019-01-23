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

Private RibbonDispatcher As PGSolutions_RibbonDispatcher.IRibbonDispatcher

Private Sub Button1_Clicked(ByVal sender As Variant)
    MsgBox "VBA CustomButton clicked.", vbOKOnly Or vbInformation, TypeName(Me)
End Sub

Friend Sub Activate()
    Dim obj As Object
    Set obj = AddInHandle.AttachButton("CustomizableButton1", Button1Strings)
    With AddInHandle
        Set Button1 = .AttachButton("CustomizableButton1", Button1Strings)
        Button1.SetImageMso "RefreshAll"
        Button1.Invalidate
        '.Invalidate
    End With
End Sub

Private Sub Class_Initialize()
    With AddInHandle
        Set Button1Strings = .NewControlStrings(Label:="This is cool!", _
                ScreenTip:="VBA-Customized Button Screentip", _
                SuperTip:="This button is completely" & vbNewLine & _
                          "customized within the VBA" & vbNewLine & _
                          "workbook.", keyTip:="", _
                AlternateLabel:="", Description:="")
    End With
End Sub

Private Function AddInHandle() As PGSolutions_RibbonDispatcher.IRibbonDispatcher
    If RibbonDispatcher Is Nothing Then Set RibbonDispatcher = NewHandle.NewBetterRibbon()
    Set AddInHandle = RibbonDispatcher
End Function

Private Function NewHandle() As PGSolutions_RibbonDispatcher.IBetterRibbon
    Set NewHandle = Application.COMAddIns("PGSolutions.BetterRibbon").Object
End Function
