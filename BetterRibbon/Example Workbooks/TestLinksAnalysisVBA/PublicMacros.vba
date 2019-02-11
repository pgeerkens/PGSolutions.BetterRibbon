Attribute VB_Name = "PublicMacros"
'''=======================================================================================
'''                Copyright (c) 2017-2019 Pieter Geerkens
'''
'''     Licensed under the MIT Licence at:
'''             https://github.com/pgeerkens/PGSolutions.BetterRibbon/blob/dev/LICENSE
'''=======================================================================================
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
