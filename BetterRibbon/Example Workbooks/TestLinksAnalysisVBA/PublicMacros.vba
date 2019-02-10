Attribute VB_Name = "PublicMacros"
Option Explicit
Option Private Module

Private Const ModuleName   As String = "PublicMacros."

Public Sub TestAll()
    On Error Resume Next
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
