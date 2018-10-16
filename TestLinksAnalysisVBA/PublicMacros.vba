Attribute VB_Name = "PublicMacros"
Option Explicit

Private Const ModuleName   As String = "PublicMacros."

''' <summary>Performs Links Analysis on ActiveWorkbook and reports the results on worksheets.</summary>
Public Sub ActiveWkbkLinks()
    On Error GoTo EH
    AddInHandle.WriteLinksAnalysisWB ActiveWorkbook
XT: Exit Sub
EH: ErrorUtils.DisplayError Err, ModuleName & "NewLinksLexer", vbOKOnly Or vbExclamation
    Resume          ' for debugging only
End Sub

Public Sub TestAll()
    RibbonUtils.TestAddinConnection
        
    LinksLexerTests.SimpleOperatorTest
    LinksLexerTests.SimpleConcatTest
    LinksLexerTests.SimpleParensTest
    LinksLexerTests.StringLiteralTest
    LinksLexerTests.ComplexRefTest
    LinksLexerTests.OpenExternRefTest

    LinksLexerTests.SimpleParseLinkTest
    LinksLexerTests.ComplexParseLinkTest
    LinksLexerTests.CellParseLinkTest
    LinksLexerTests.ArrayNamedRangeTest
End Sub
