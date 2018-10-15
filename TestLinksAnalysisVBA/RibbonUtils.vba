Attribute VB_Name = "RibbonUtils"
Option Explicit
Option Private Module
Private Const ModuleName As String = "RibbonUtilities."

Public Function NewLinksLexer(CellRef As ISourceCellRef, Formula As String) As LinksAnalyzer2.ILinksLexer
    On Error GoTo EH
    With AddInHandle
        Set NewLinksLexer = .NewLinksLexer(CellRef, Formula)
    End With
XT: Exit Function
EH: ErrorUtils.ReRaiseError Err, ModuleName & "NewLinksLexer"
    Resume          ' for debugging only
End Function

Public Function NewCellRef() As LinksAnalyzer2.ISourceCellRef
    On Error GoTo EH
    With AddInHandle
        Set NewCellRef = .NewSourceCellRef(ThisWorkbook, "MyTab", "A1")
    End With
XT: Exit Function
EH: ErrorUtils.ReRaiseError Err, ModuleName & "NewCellRef"
    Resume          ' for debugging only
End Function

Public Function AddInHandle() As LinksAnalyzer2.ILinksAnalyzer
    On Error GoTo EH
    Set AddInHandle = Application.COMAddIns("ExcelRibbon").Object
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
    
        StepName = "Get SourceCellRef to Object"
        Set obj = .NewSourceCellRef(ThisWorkbook, "TaName", "CellRef")
        Messages = Messages & StepName & " - success" & vbNewLine
        
        StepName = "Get SourceCellRef to ISourceCellRef"
        Dim CellRef As ISourceCellRef
        Set CellRef = .NewSourceCellRef(ThisWorkbook, "TaName", "CellRef")
        Messages = Messages & StepName & " - success" & vbNewLine
    
        StepName = "Get LinksLexer to Object"
        Set obj = .NewLinksLexer(CellRef, "Formula")
        Messages = Messages & StepName & " - success" & vbNewLine
        
        StepName = "Get LinksLexer to Object"
        Dim Lexer As ILinksLexer
        Set Lexer = .NewLinksLexer(CellRef, "Formula")
        Messages = Messages & StepName & " - success" & vbNewLine
    End With
    MsgBox Messages, vbOKOnly, "TestAddinConnection"
    
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, StepName
    Resume XT
    Resume          ' for debugging only
End Sub
