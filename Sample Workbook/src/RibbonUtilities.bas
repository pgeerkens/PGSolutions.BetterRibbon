Attribute VB_Name = "RibbonUtilities"
Option Explicit
Option Private Module

Private Const MModuleName   As String = "RibbonUtilities."

Public Enum ControlSize
    rdRegular = RdControlSize_rdRegular
    rdLarge = RdControlSize_rdLarge
End Enum

''' <summary>EventHandler for RibbonLoad, initializing the ViewModel and Model.</summary>
Public Const MsgBoxTitle    As String = "PGSolutions Sample Ribbon"

''' <summary>EventHandler for RibbonLoad, initializing the ViewModel and Model.</summary>
Public Sub DefaultButtonAction(ByVal Button As RibbonButton)
    MsgBox Button.Label & " Pressed", vbOKOnly Or vbInformation, MsgBoxTitle
End Sub

''' <summary>EventHandler for RibbonLoad, initializing the ViewModel and Model.</summary>
''' <param name="IsLarge">BitFlag for desired presentation of Image & Label.</param>
Public Sub SetButtonView(ByVal SelectedIndex As Integer, ParamArray Buttons())
    Dim i As Long
    Dim ShowLabel As Boolean: ShowLabel = ((SelectedIndex + 1) And 1) <> 0
    Dim ShowImage As Boolean: ShowImage = ((SelectedIndex + 1) And 2) <> 0
    For i = LBound(Buttons) To UBound(Buttons)
        Buttons(i).ShowLabel = ShowLabel
        Buttons(i).ShowImage = ShowImage
    Next i
End Sub

''' <summary>Sets the size of all supplied Buttons, and returns the value of IsLarge.</summary>
''' <param name="IsLarge">True if Buttons should be resized Large; else false.</param>
''' <param name="Buttons">The ParamArray of Buttons to be resized.</param>
Public Function ToggleCustomSize(ByVal IsLarge As Boolean, ParamArray Buttons()) As Boolean
    On Error GoTo EH
    Dim i As Long
    For i = LBound(Buttons) To UBound(Buttons)
        Buttons(i).Size = IIf(IsLarge, rdLarge, rdRegular)
    Next i
    ToggleCustomSize = Not IsLarge
XT: Exit Function
EH: ReraiseError Err, ModuleName & "ToggleCustomSize"
    Resume      ' for debugging only
End Function
