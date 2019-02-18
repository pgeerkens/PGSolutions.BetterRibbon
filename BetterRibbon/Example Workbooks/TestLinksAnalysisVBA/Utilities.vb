Attribute VB_Name = "Utilities"
'''=======================================================================================
'''                Copyright (c) 2017-2019 Pieter Geerkens
'''
'''     Licensed under the MIT Licence at:
'''             https://github.com/pgeerkens/PGSolutions.BetterRibbon/blob/dev/LICENSE
'''=======================================================================================
Option Explicit
Option Private Module
Private Const ModuleName As String = "RibbonUtils."

Public Function NewLinksLexer(CellRef As ISourceCellRef, Formula As String) As ILinksLexer
    On Error GoTo EH
    With AddInHandle
        Set NewLinksLexer = .NewLinksLexer(CellRef, Formula)
    End With
XT: Exit Function
EH: ErrorUtils.ReRaiseError Err, ModuleName & "NewLinksLexer"
    Resume          ' for debugging only
End Function

Public Property Get DummyCellRef() As ISourceCellRef
    Set DummyCellRef = New SourceCellRef
End Property

Public Function AddInHandle() As ILinksAnalyzer
    On Error GoTo EH
    Set AddInHandle = Application.COMAddIns("PGSolutions.BetterRibbon").Object.NewLinksAnalyzer
XT: Exit Function
EH: ErrorUtils.ReRaiseError Err, ModuleName & "AddInHandle"
    Resume          ' for debugging only
End Function

Public Function TestAddinConnection() As Boolean
    On Error GoTo EH
    TestAddinConnection = False
    Dim StepName    As String, _
        objAddIns   As Object, _
        objHandle   As Object, _
        Handle      As ILinksAnalyzer, _
        objLexer    As Object, _
        Lexer       As ILinksLexer, _
        Messages    As String
    
    StepName = "Get COMAddIns as Object"
    Set objAddIns = Application.COMAddIns("PGSolutions.BetterRibbon")
    Messages = Messages & vbNewLine & "Success - " & StepName
    
    StepName = "Get AddInHandle as Object"
    Set objHandle = objAddIns.Object
    Messages = Messages & vbNewLine & "Success - " & StepName
    
    StepName = "Get AddInHandle as ILinksAnalyzer"
    Set Handle = AddInHandle
    Messages = Messages & vbNewLine & "Success - " & StepName
    
    With Handle
        Messages = Messages & vbNewLine & "Success - " & StepName
        
        StepName = "Get LinksLexer to Object"
        Set objLexer = .NewLinksLexer(DummyCellRef, "Formula")
        Messages = Messages & vbNewLine & "Success - " & StepName
        
        StepName = "Get LinksLexer to ILinksLexer"
        Set Lexer = .NewLinksLexer(DummyCellRef, "Formula")
        Messages = Messages & vbNewLine & "Success - " & StepName
    End With
    TestAddinConnection = True
    
XT: MsgBox Messages, vbOKOnly, "TestAddinConnection"
    Exit Function
EH: ErrorUtils.ReRaiseError Err, ModuleName & "TestAddinConnection"
    Resume XT
    Resume          ' for debugging only
End Function
