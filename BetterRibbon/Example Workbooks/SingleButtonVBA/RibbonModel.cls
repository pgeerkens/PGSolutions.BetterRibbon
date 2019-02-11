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

Private MRibbonDispatcher   As PGSolutions_RibbonDispatcher.IRibbonDispatcher

Private Button1Strings      As PGSolutions_RibbonDispatcher.IRibbonControlStrings
Private WithEvents Button1  As PGSolutions_RibbonDispatcher.RibbonButtonModel
Attribute Button1.VB_VarHelpID = -1

Private Sub Button1_Clicked(ByVal sender As Variant)
    ButtonProcessing.Button1_Processing
End Sub

Friend Sub Activate()
    Button1.Attach "CustomizableButton1", Button1Strings
    Button1.Invalidate
    RibbonDispatcher.InvalidateCustomControlsGroup
End Sub

Private Function RibbonDispatcher() As PGSolutions_RibbonDispatcher.IRibbonDispatcher
    If MRibbonDispatcher Is Nothing Then
        Set MRibbonDispatcher = Application.COMAddIns("PGSolutions.BetterRibbon").Object.NewBetterRibbon()
    End If
    Set RibbonDispatcher = MRibbonDispatcher
End Function

Private Sub Class_Initialize()
    With RibbonDispatcher
        Set Button1Strings = .NewControlStrings(Label:="This is cool!", _
                ScreenTip:="VBA-Customized Button Screentip", _
                SuperTip:="This button is completely" & vbNewLine & _
                          "customized within the VBA" & vbNewLine & _
                          "workbook.", keyTip:="", _
                AlternateLabel:="", Description:="")
                
        Set Button1 = .NewRibbonButtonModel:  Button1.SetImageMso "MacroPlay": Button1.Invalidate
    End With
End Sub
