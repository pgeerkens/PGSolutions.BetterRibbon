Attribute VB_Name = "PublicMacros"
Option Explicit

Private Const ModuleName   As String = "PublicMacros."

Public Sub TestAll()
    If RibbonUtils.TestAddinConnection Then
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
    End If
End Sub
