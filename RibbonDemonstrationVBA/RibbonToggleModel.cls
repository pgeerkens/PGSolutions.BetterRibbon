VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RibbonToggleModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private MIsPressed              As Boolean
Private WithEvents MViewModel   As PGSolutions_RibbonDispatcher.RibbonToggleButton
Attribute MViewModel.VB_VarHelpID = -1

Public Event Toggled(ByVal Sender As Variant, ByVal IsPressed As Boolean)

Implements IBooleanSource

Friend Function IBooleanSource_Getter() As Boolean
    IBooleanSource_Getter = MIsPressed
End Function

Private Sub MViewModel_Toggled(ByVal Sender As Variant, ByVal IsPressed As Boolean)
    On Error GoTo EH
    MIsPressed = IsPressed
    RaiseEvent Toggled(Sender, IsPressed)
    MViewModel.Invalidate
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, TypeName(Me) & "MViewModel_Toggled", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub

Friend Property Get ViewModel() As PGSolutions_RibbonDispatcher.RibbonToggleButton
    Set ViewModel = MViewModel
End Property

Friend Property Get IsPressed() As Boolean
    IsPressed = MIsPressed
End Property
Friend Property Let IsPressed(ByVal IsPressed As Boolean)
    MIsPressed = IsPressed
    MViewModel.Invalidate
End Property

Friend Function Attach(ByVal ControlId As String, _
        ByVal Strings As PGSolutions_RibbonDispatcher.IRibbonControlStrings _
) As RibbonToggleModel
    On Error GoTo EH
    Set MViewModel = AddInHandle.AttachToggle(ControlId, Strings, Me)
    Set Attach = Me
XT: Exit Function
EH: ErrorUtils.ReRaiseError Err, TypeName(Me) & "Activate"
    Resume          ' for debugging only
End Function
