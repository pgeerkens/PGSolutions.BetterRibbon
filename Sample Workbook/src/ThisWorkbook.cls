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

''' <summary>Returns whether or not to display popups during load, for debugging ease.</summary>
Private Const ShowAlerts     As Boolean = True

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''   DON'T EDIT THIS MODULE   '''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private MRibbonModel    As IRibbonModel

Implements IRibbonWorkbook
 
Private Property Get IRibbonWorkbook_ViewModel() As RibbonViewModel
    Set IRibbonWorkbook_ViewModel = MRibbonModel.RibbonViewModel
End Property

Private Function IRibbonWorkbook_InitializeRibbonModel(ByVal ribbonUI As Office.IRibbonUI) As IRibbonModel
    With New RibbonModel
        Set MRibbonModel = .Initialize(ribbonUI)
        Set IRibbonWorkbook_InitializeRibbonModel = MRibbonModel
    End With
    MRibbonModel.ActivateTab    ' Must follow the assignment aobve - for obvious reasons
End Function

Private Property Get IRibbonWorkbook_Path() As String
    IRibbonWorkbook_Path = ThisWorkbook.Path
End Property

Private Sub Workbook_Open()
    On Error GoTo EH
    If ShowAlerts Then DisplayAlert "Workbook_Open"
XT: Exit Sub
EH: DisplayError Err, TypeName(Me) & "Workbook_Open"
    Resume XT
    Resume      ' for debugging only
End Sub
