Attribute VB_Name = "RibbonUtils"
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
    Set AddInHandle = Application.COMAddIns("BetterRibbon").Object
XT: Exit Function
EH: ErrorUtils.ReRaiseError Err, ModuleName & "AddInHandle"
    Resume          ' for debugging only
End Function

Public Sub TestAddinConnection()
    On Error GoTo EH
    Dim StepName    As String, _
        obj         As Object, _
        Messages    As String
        
    StepName = "Get AddInHandle"
    With AddInHandle
    Messages = Messages & StepName & " - success" & vbNewLine
    
        StepName = "Get LinksLexer to Object"
        Set obj = .NewLinksLexer(DummyCellRef, "Formula")
        Messages = Messages & StepName & " - success" & vbNewLine
        
        StepName = "Get LinksLexer to Object"
        Dim Lexer As ILinksLexer
        Set Lexer = .NewLinksLexer(DummyCellRef, "Formula")
        Messages = Messages & StepName & " - success" & vbNewLine
    End With
    MsgBox Messages, vbOKOnly, "TestAddinConnection"
    
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, StepName
    Resume XT
    Resume          ' for debugging only
End Sub
