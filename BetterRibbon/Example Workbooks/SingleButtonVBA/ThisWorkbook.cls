VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'''=======================================================================================
'''                Copyright (c) 2017-2019 Pieter Geerkens
'''
'''     Licensed under the MIT Licence at:
'''             https://github.com/pgeerkens/PGSolutions.BetterRibbon/blob/dev/LICENSE
'''=======================================================================================
Option Explicit
Private Const COMAddInName  As String = "PGSolutions.BetterRibbon"
Private MModelServer        As PGSolutions_RibbonDispatcher.IModelServer
Private MRibbonModel        As RibbonModel

Public Function ModelServer() As PGSolutions_RibbonDispatcher.IModelServer
    If MModelServer Is Nothing Then
        Set MModelServer = Application.COMAddIns(COMAddInName).Object _
                .NewBetterRibbon(New ResourceLoader)
    End If
    Set ModelServer = MModelServer
End Function

Public Sub Register()
    Application.COMAddIns(COMAddInName).Object.RegisterWorkbook ThisWorkbook.Name
End Sub

Private Sub Workbook_Activate()
    If MRibbonModel Is Nothing Then Set MRibbonModel = New RibbonModel
End Sub

Private Sub Workbook_Open()
    If DeskTop(False) = "D:\Users\Pieter\Desktop\" _
    Or ThisWorkbook.Path = DeskTop(True) & "Example Workbooks" _
    Or ThisWorkbook.Path = DeskTop(False) & "Example Workbooks" Then _
        MsgBox "Pause for Ctrl-Break to ease debugging." & vbNewLine & vbNewLine & _
               "This message can be disabled by moving the workbook" & vbNewLine & _
               "out of the Desktop folder 'Example Workbooks'.", _
               vbOKOnly, ThisWorkbook.Name
End Sub
