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

Private Sub Workbook_Open()
    On Error GoTo EH
    If ShowAlerts Then DisplayAlert "Workbook_Open"
XT: Exit Sub
EH: DisplayError Err, TypeName(Me) & "Workbook_Open"
    Resume XT
    Resume      ' for debugging only
End Sub
