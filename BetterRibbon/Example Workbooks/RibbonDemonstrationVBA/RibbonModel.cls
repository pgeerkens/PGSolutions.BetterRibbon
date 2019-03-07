VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RibbonModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''=======================================================================================
'''                Copyright (c) 2017-2019 Pieter Geerkens
'''
'''     Licensed under the MIT Licence at:
'''             https://github.com/pgeerkens/PGSolutions.BetterRibbon/blob/dev/LICENSE
'''=======================================================================================
Option Explicit
Private CustomGroup          As IGroupModel

Private WithEvents Toggle1   As ToggleModel
Attribute Toggle1.VB_VarHelpID = -1
Private WithEvents Toggle2   As ToggleModel
Attribute Toggle2.VB_VarHelpID = -1
Private WithEvents Toggle3   As ToggleModel
Attribute Toggle3.VB_VarHelpID = -1
Private WithEvents Toggle4   As ToggleModel
Attribute Toggle4.VB_VarHelpID = -1

Private WithEvents DropDown1 As DropDownModel
Attribute DropDown1.VB_VarHelpID = -1
Private WithEvents DropDown2 As DropDownModel
Attribute DropDown2.VB_VarHelpID = -1
Private WithEvents Dropdown3 As DropDownModel
Attribute Dropdown3.VB_VarHelpID = -1

Private WithEvents Button1   As ButtonModel
Attribute Button1.VB_VarHelpID = -1
Private WithEvents Button2   As ButtonModel
Attribute Button2.VB_VarHelpID = -1
Private WithEvents Button3   As ButtonModel
Attribute Button3.VB_VarHelpID = -1

Private Sub Toggle1_Toggled(ByVal control As IRibbonControl, ByVal IsPressed As Boolean)
    On Error GoTo EH
    CustomGroup.SetShowInactive IsPressed
    Toggle1.SetImage ThisWorkbook.ModelServer.NewImageObjectMso(ToggleImage(IsPressed))
    
    Invalidate
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, TypeName(Me) & ".Toggle1_Toggled", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub

Private Sub Toggle2_Toggled(ByVal control As IRibbonControl, ByVal IsPressed As Boolean)
    On Error GoTo EH
    Toggle1.IsLarge = Not Toggle2.IsPressed
    Toggle2.IsLarge = Not Toggle2.IsPressed
    Toggle3.IsLarge = Not Toggle2.IsPressed
    Toggle2.SetImage ThisWorkbook.ModelServer.NewImageObjectMso(ToggleImage(Toggle2.IsPressed))
    
    Invalidate
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, TypeName(Me) & ".Toggle2_Toggled", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub

Private Sub Toggle3_Toggled(ByVal control As IRibbonControl, ByVal IsPressed As Boolean)
    On Error GoTo EH
    Button1.IsLarge = Not IsPressed
    Button2.IsLarge = Not IsPressed
    Button3.IsLarge = Not IsPressed
    Toggle3.SetImage ThisWorkbook.ModelServer.NewImageObjectMso(ToggleImage(IsPressed))
    DropDown1.IsEnabled = Toggle3.IsPressed
    
    Invalidate
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, TypeName(Me) & ".Toggle3_Toggled", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub

Private Sub Toggle4_Toggled(ByVal control As IRibbonControl, ByVal IsPressed As Boolean)
    On Error GoTo EH
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Control Strings")
    ws.Visible = IIf(ws.Visible = xlSheetVeryHidden, xlSheetVisible, xlSheetVeryHidden)
    Toggle4.IsPressed = (ws.Visible = xlSheetVisible)
    ws.Activate
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, TypeName(Me) & ".Toggle3_Toggled", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub

Private Sub DropDown1_SelectionMade(ByVal control As Office.IRibbonControl, ByVal selectedId As String, ByVal SelectedIndex As Long)
    On Error GoTo EH
    SetImageAndLabel SelectedIndex, Button1, Button2, Button3
    
    Invalidate
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, TypeName(Me) & ".DropDown1_SelectionMade", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub

Private Sub DropDown2_SelectionMade(ByVal control As Office.IRibbonControl, ByVal selectedId As String, ByVal SelectedIndex As Long)
    On Error GoTo EH
    MsgBox "DropDown2 has not been implemented yet", vbOKOnly Or vbInformation, TypeName(Me)
    Invalidate
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, TypeName(Me) & ".DropDown2_SelectionMade", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub

Private Sub DropDown3_SelectionMade(ByVal control As Office.IRibbonControl, ByVal selectedId As String, ByVal SelectedIndex As Long)
    On Error GoTo EH
    MsgBox "DropDown2 has not been implemented yet", vbOKOnly Or vbInformation, TypeName(Me)
    Invalidate
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, TypeName(Me) & ".DropDown3_SelectionMade", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub

Private Sub Button1_Clicked(ByVal control As IRibbonControl)
    On Error GoTo EH
    Static IsToggle As Boolean
    IsToggle = AlternateToggle(ThisWorkbook.ModelServer, IsToggle, Toggle2, "CustomVbaToggle2", "CustomVbaCheckBox2")
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, TypeName(Me) & ".Button1_Clicked", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub

Private Sub Button2_Clicked(ByVal control As IRibbonControl)
    On Error GoTo EH
    Static IsToggle As Boolean
    IsToggle = AlternateToggle(ThisWorkbook.ModelServer, IsToggle, Toggle3, "CustomVbaToggle3", "CustomVbaCheckBox3")
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, TypeName(Me) & ".Button2_Clicked", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub

Private Sub Button3_Clicked(ByVal control As IRibbonControl)
    On Error GoTo EH
    Class_Initialize
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, TypeName(Me) & ".Button3_Clicked", vbOKOnly Or vbInformation
    Resume XT
    Resume          ' for debugging only
End Sub

Private Sub Invalidate()
    On Error Resume Next
    Toggle4.IsPressed = (ThisWorkbook.Worksheets("Control Strings").Visible = xlSheetVisible)
    DropDown1.IsEnabled = Toggle3.IsPressed
    CustomGroup.SetShowInactive Toggle1.IsPressed
    
    Toggle1.Invalidate
    Toggle2.Invalidate
    Toggle3.Invalidate
    Toggle4.Invalidate
    
    DropDown1.Invalidate
    DropDown2.Invalidate
    Dropdown3.Invalidate
    
    Button1.Invalidate
    Button2.Invalidate
    Button3.Invalidate
    
    CustomGroup.Invalidate
    Application.StatusBar = "Ready ...'"
XT: Exit Sub
End Sub

Private Sub Class_Initialize()
    On Error GoTo EH
    If Not CustomGroup Is Nothing Then CustomGroup.Detach
    With ThisWorkbook.ModelServer
        Set CustomGroup = .GetGroupModel("CustomizableGroup")
    End With
    
    ' Demonstrate how the models can treat the viewmodels as a pure 'skin'
    With ThisWorkbook.ModelFactory
        Set Toggle1 = .NewToggleModel("Toggle1").Attach("CustomVbaToggle1") _
                      .SetImage(.NewImageObjectMso("MarginsShowHide"))
        Set Toggle2 = .NewToggleModel("Toggle2").Attach("CustomVbaCheckBox2") _
                      .SetImage(.NewImageObjectMso("MarginsShowHide"))
        Toggle2.IsPressed = True
        Set Toggle3 = .NewToggleModel("Toggle3").Attach("CustomVbaCheckBox3") _
                      .SetImage(.NewImageObjectMso("MarginsShowHide"))
        Set Toggle4 = .NewToggleModel("Toggle4").Attach("CustomVbaCheckBox1") _
                      .SetImage(.NewImageObjectMso("MarginsShowHide"))
        
        Set DropDown1 = .NewDropDownModel("DropDown1").Attach("CustomVbaDropDown1") _
                        .SetImage(.NewImageObjectMso("MacroPlay")) _
                        .AddSelectableModel(.NewSelectableModel("LabelOnly")) _
                        .AddSelectableModel(.NewSelectableModel("ImageOnly")) _
                        .AddSelectableModel(.NewSelectableModel("LabelAndImage"))
        DropDown1.SelectedIndex = 2
        DropDown1.ShowImage = True
        
        Set Button1 = .NewButtonModel("Button1").Attach("CustomizableButton1") _
                      .SetImage(.NewImageObjectMso("MacroPlay"))
        Set Button2 = .NewButtonModel("Button2").Attach("CustomizableButton2") _
                      .SetImage(.NewImageObjectMso("MacroPlay"))
        Set Button3 = .NewButtonModel("Button3").Attach("CustomizableButton3") _
                      .SetImage(.NewImageObjectMso("Refresh"))
    End With
       
    Invalidate
XT: Exit Sub
EH: Select Case MsgBoxAbortRetryIgnore(Err, TypeName(Me) & ".Class_Initialize")
        Case vbAbort:  ErrorUtils.ReRaiseError Err, TypeName(Me) & ".Class_Initialize"
        Case vbRetry:  Resume
        Case vbIgnore: Resume Next
    End Select
    Resume          ' for debugging only
End Sub
