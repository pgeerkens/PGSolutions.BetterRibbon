Attribute VB_Name = "LinksLexerTests"
'''=======================================================================================
'''                Copyright (c) 2017-2019 Pieter Geerkens
'''
'''     Licensed under the MIT Licence at:
'''             https://github.com/pgeerkens/PGSolutions.BetterRibbon/blob/dev/LICENSE
'''=======================================================================================
Option Explicit
Option Private Module

Private Const mModuleName As String = "LinksLexerTests."

Public Sub SimpleOperatorTest()
    Const MethodName As String = mModuleName & "SimpleOperatorTest"
    
    On Error GoTo EH
    Const Formula As String = "4 + 5"
    
    Dim Lexer As ILinksLexer: Set Lexer = NewLinksLexer(DummyCellRef, Formula)
        ScanCheck MethodName, Lexer, EToken_Number, "4"
        ScanCheck MethodName, Lexer, EToken_UnaryOperator, "+"
        ScanCheck MethodName, Lexer, EToken_Number, "5"
        ScanCheckEOT MethodName, Lexer
    MsgBox "Successfully scanned: " & vbNewLine & Formula, vbOKOnly, MethodName
XT: Exit Sub
EH: Select Case MsgBoxAbortRetryIgnore(Err, mModuleName & "SimpleOperatorTest")
        Case vbRetry:  Resume
        Case vbIgnore: Resume Next
    End Select
    Resume XT
    Resume
End Sub

Public Sub SimpleConcatTest()
    Const MethodName As String = mModuleName & "SimpleConcatTest"
    
    On Error GoTo EH
    Const Formula As String = "B4&"" YTD"""
    
    Dim Lexer As ILinksLexer: Set Lexer = NewLinksLexer(DummyCellRef, Formula)
        ScanCheck MethodName, Lexer, EToken_Identifier, "B4"
        ScanCheck MethodName, Lexer, EToken_BinaryOperator, "&"
        ScanCheck MethodName, Lexer, EToken_StringLiteral, """ YTD"""
        ScanCheckEOT MethodName, Lexer
    MsgBox "Successfully scanned: " & vbNewLine & Formula, vbOKOnly, MethodName
XT: Exit Sub
EH: Select Case MsgBoxAbortRetryIgnore(Err, MethodName)
        Case vbRetry:  Resume
        Case vbIgnore: Resume Next
    End Select
    Resume XT
    Resume
End Sub

Public Sub SimpleParensTest()
    Const MethodName As String = mModuleName & "SimpleParensTest"
    
    On Error GoTo EH
    Const Formula As String = "(4+5)"
    
    Dim Lexer As ILinksLexer: Set Lexer = NewLinksLexer(DummyCellRef, Formula)
        ScanCheck MethodName, Lexer, EToken_OpenParen, "("
        ScanCheck MethodName, Lexer, EToken_Number, "4"
        ScanCheck MethodName, Lexer, EToken_UnaryOperator, "+"
        ScanCheck MethodName, Lexer, EToken_Number, "5"
        ScanCheck MethodName, Lexer, EToken_CloseParen, ")"
        ScanCheckEOT MethodName, Lexer
    MsgBox "Successfully scanned: " & vbNewLine & Formula, vbOKOnly, MethodName
XT: Exit Sub
EH: Select Case MsgBoxAbortRetryIgnore(Err, MethodName)
        Case vbRetry:  Resume
        Case vbIgnore: Resume Next
    End Select
    Resume XT
    Resume
End Sub

Public Sub StringLiteralTest()
    Const MethodName As String = mModuleName & "StringLiteralTest"
    
    On Error GoTo EH
    Const Formula As String = "=MID(C2, FIND("" '"", C2, 1)+1,  FIND(""]"", C2, 1)-FIND(""'"", C2, 1))"
    
    Dim Lexer As ILinksLexer: Set Lexer = NewLinksLexer(DummyCellRef, Formula)
        ScanCheck MethodName, Lexer, EToken_Equals, "="
        ScanCheck MethodName, Lexer, EToken_Identifier, "MID"
        ScanCheck MethodName, Lexer, EToken_OpenParen, "("
        ScanCheck MethodName, Lexer, EToken_Identifier, "C2"
        ScanCheck MethodName, Lexer, EToken_Comma, ","
        ScanCheck MethodName, Lexer, EToken_Identifier, "FIND"
        ScanCheck MethodName, Lexer, EToken_OpenParen, "("
        ScanCheck MethodName, Lexer, EToken_StringLiteral, """ '"""
        ScanCheck MethodName, Lexer, EToken_Comma, ","
        ScanCheck MethodName, Lexer, EToken_Identifier, "C2"
        
        ScanCheck MethodName, Lexer, EToken_Comma, ","
        ScanCheck MethodName, Lexer, EToken_Number, "1"
        ScanCheck MethodName, Lexer, EToken_CloseParen, ")"
        ScanCheck MethodName, Lexer, EToken_UnaryOperator, "+"
        ScanCheck MethodName, Lexer, EToken_Number, "1"
        ScanCheck MethodName, Lexer, EToken_Comma, ","
        ScanCheck MethodName, Lexer, EToken_Identifier, "FIND"
        ScanCheck MethodName, Lexer, EToken_OpenParen, "("
        ScanCheck MethodName, Lexer, EToken_StringLiteral, """]"""
        ScanCheck MethodName, Lexer, EToken_Comma, ","
        
        ScanCheck MethodName, Lexer, EToken_Identifier, "C2"
        ScanCheck MethodName, Lexer, EToken_Comma, ","
        ScanCheck MethodName, Lexer, EToken_Number, "1"
        ScanCheck MethodName, Lexer, EToken_CloseParen, ")"
        ScanCheck MethodName, Lexer, EToken_UnaryOperator, "-"
        ScanCheck MethodName, Lexer, EToken_Identifier, "FIND"
        ScanCheck MethodName, Lexer, EToken_OpenParen, "("
        ScanCheck MethodName, Lexer, EToken_StringLiteral, """'"""
        ScanCheck MethodName, Lexer, EToken_Comma, ","
        ScanCheck MethodName, Lexer, EToken_Identifier, "C2"
        
        ScanCheck MethodName, Lexer, EToken_Comma, ","
        ScanCheck MethodName, Lexer, EToken_Number, "1"
        ScanCheck MethodName, Lexer, EToken_CloseParen, ")"
        ScanCheck MethodName, Lexer, EToken_CloseParen, ")"
        
        ScanCheckEOT MethodName, Lexer
    MsgBox "Successfully scanned: " & vbNewLine & Formula, vbOKOnly, MethodName
XT: Exit Sub
EH: Select Case MsgBoxAbortRetryIgnore(Err, MethodName)
        Case vbRetry:  Resume
        Case vbIgnore: Resume Next
    End Select
    Resume XT
    Resume
End Sub

Public Sub ComplexRefTest()
    Const MethodName As String = mModuleName & "ComplexRefTest"
    
    On Error GoTo EH
    Const Formula As String = _
        "=VLOOKUP(A18,'G:\can\Income Stmt Mapping\[IS Mapping.xlsx]IS_line names'!$A$6:$B$400,2,FALSE)"
    
    Dim Lexer As ILinksLexer: Set Lexer = NewLinksLexer(DummyCellRef, Formula)
        ScanCheck MethodName, Lexer, EToken_Equals, "="
        ScanCheck MethodName, Lexer, EToken_Identifier, "VLOOKUP"
        ScanCheck MethodName, Lexer, EToken_OpenParen, "("
        ScanCheck MethodName, Lexer, EToken_Identifier, "A18"
        ScanCheck MethodName, Lexer, EToken_Comma, ","
        
        ScanCheck MethodName, Lexer, EToken_ExternRef, _
            "'G:\can\Income Stmt Mapping\[IS Mapping.xlsx]IS_line names'"
        ScanCheck MethodName, Lexer, EToken_Bang, "!"
        ScanCheck MethodName, Lexer, EToken_Identifier, "$A$6:$B$400"
        ScanCheck MethodName, Lexer, EToken_Comma, ","
        
        ScanCheck MethodName, Lexer, EToken_Number, "2"
        ScanCheck MethodName, Lexer, EToken_Comma, ","
        ScanCheck MethodName, Lexer, EToken_Identifier, "FALSE"
        ScanCheck MethodName, Lexer, EToken_CloseParen, ")"
        
        ScanCheckEOT MethodName, Lexer
    MsgBox "Successfully scanned+: " & vbNewLine & Formula, vbOKOnly, MethodName
XT: Exit Sub
EH: Select Case MsgBoxAbortRetryIgnore(Err, MethodName)
        Case vbRetry:  Resume
        Case vbIgnore: Resume Next
    End Select
    Resume XT
    Resume
End Sub

Public Sub OpenExternRefTest()
    Const MethodName As String = mModuleName & "OpenExternRefTest"
    
    On Error GoTo EH
    Const Formula As String = "=[RibbonDemonstration.xlsb]Sheet1!$A$2+1"
    Dim ExtLinks As IExternalLinks
    Set ExtLinks = AddInHandle.Parse(DummyCellRef, Formula).Links
    With ExtLinks.Item(0)
        If .TargetPath <> "open workbook w/o a path" Then _
             Err.Raise 1, MethodName, "Incorrect Path found"
        
        If .TargetFile <> "RibbonDemonstration.xlsb" Then _
            Err.Raise 1, MethodName, "Incorrect FileName found"

        If .TargetTab <> "Sheet1" Then _
            Err.Raise 1, MethodName, "Incorrect TabName found"

        If .TargetCell <> "$A$2" Then _
            Err.Raise 1, MethodName, "Incorrect Cell found"
    End With
    MsgBox "Successfully parsed: " & _
        vbNewLine & Formula & "as" & _
        vbNewLine & _
        vbNewLine & "Path: " & ExtLinks.Item(0).TargetPath, vbOKOnly, MethodName
    
XT: Exit Sub
    
EH: Select Case MsgBoxAbortRetryIgnore(Err, MethodName)
        Case vbRetry:  Resume
        Case vbIgnore: Resume Next
    End Select
    Resume XT
    Resume
End Sub

Public Sub SimpleParseLinkTest()
    Const MethodName As String = mModuleName & "SimpleParseLinkTest"
    
    On Error GoTo EH
    Const Formula As String = _
        "=VLOOKUP(A18,'S:\can\Affinity\actuar\SPONSOR\VALN\Income Statement Mapping\[IS Mapping.xlsx]IS_line names'!$A$6:$B$400,2,FALSE)"
                
    Dim ExtLinks As IExternalLinks
    Set ExtLinks = AddInHandle.Parse(DummyCellRef, Formula).Links
    With ExtLinks.Item(0)
        If .TargetPath <> "S:\can\Affinity\actuar\SPONSOR\VALN\Income Statement Mapping\" Then _
             Err.Raise 1, MethodName, "Incorrect Path found"
        
        If .TargetFile <> "IS Mapping.xlsx" Then _
            Err.Raise 1, MethodName, "Incorrect FileName found"

        If .TargetTab <> "IS_line names" Then _
            Err.Raise 1, MethodName, "Incorrect TabName found"

        If .TargetCell <> "$A$6:$B$400" Then _
            Err.Raise 1, MethodName, "Incorrect Cell found"
    End With
    MsgBox "Successfully parsed: " & _
        vbNewLine & Formula & "as" & _
        vbNewLine & _
        vbNewLine & "Path: " & ExtLinks.Item(0).TargetPath, vbOKOnly, MethodName
    
XT: Exit Sub
    
EH: Select Case MsgBoxAbortRetryIgnore(Err, MethodName)
        Case vbRetry:  Resume
        Case vbIgnore: Resume Next
    End Select
    Resume XT
    Resume
End Sub

Public Sub ComplexParseLinkTest()
    Const MethodName As String = mModuleName & "ComplexParseLinkTest"
    Const PathPrefix As String = "S:\can\Finance\actuarial\ASSC\Institutional\Reporting\2016\"
    Const Formula As String = _
        "=SUM('S:\can\Finance\actuarial\ASSC\Institutional\Reporting\2016\M03\Reserves\CRR\Affinity\[CCPE Experience Report_2016 Q1.xls]INTERIM REPORT'!$V$16," & vbNewLine & _
        "     'S:\can\Finance\actuarial\ASSC\Institutional\Reporting\2016\M03\Reserves\CRR\Affinity\[CCPE Experience Report_2016 Q1.xls]INTERIM REPORT'!$W$16)" & vbNewLine & _
        "-SUM('S:\can\finance\actuarial\ASSC\Institutional\Reporting\2015\M12\Reserves\CRR\Affinity\[CCPE Experience Report_2015 Q4 - Corrected.xls]INTERIM REPORT'!$V$16," & vbNewLine & _
        "     'S:\can\finance\actuarial\ASSC\Institutional\Reporting\2015\M12\Reserves\CRR\Affinity\[CCPE Experience Report_2015 Q4 - Corrected.xls]INTERIM REPORT'!$W$16)" & vbNewLine
                
    On Error GoTo EH
    Dim ExtLinks As IExternalLinks
    Set ExtLinks = AddInHandle.Parse(DummyCellRef, Formula).Links
    With ExtLinks
        If .Item(0).TargetPath <> "S:\can\Finance\actuarial\ASSC\Institutional\Reporting\2016\M03\Reserves\CRR\Affinity\" _
        Or .Item(1).TargetPath <> "S:\can\Finance\actuarial\ASSC\Institutional\Reporting\2016\M03\Reserves\CRR\Affinity\" _
        Or .Item(2).TargetPath <> "S:\can\finance\actuarial\ASSC\Institutional\Reporting\2015\M12\Reserves\CRR\Affinity\" _
        Or .Item(3).TargetPath <> "S:\can\finance\actuarial\ASSC\Institutional\Reporting\2015\M12\Reserves\CRR\Affinity\" _
        Then Err.Raise 1, MethodName, "Incorrect Path found"
        
        If .Item(0).TargetFile <> "CCPE Experience Report_2016 Q1.xls" _
        Or .Item(1).TargetFile <> "CCPE Experience Report_2016 Q1.xls" _
        Or .Item(2).TargetFile <> "CCPE Experience Report_2015 Q4 - Corrected.xls" _
        Or .Item(3).TargetFile <> "CCPE Experience Report_2015 Q4 - Corrected.xls" _
        Then Err.Raise 1, MethodName, "Incorrect FileName found"

        If .Item(0).TargetTab <> "INTERIM REPORT" _
        Or .Item(1).TargetTab <> "INTERIM REPORT" _
        Or .Item(2).TargetTab <> "INTERIM REPORT" _
        Or .Item(3).TargetTab <> "INTERIM REPORT" _
        Then Err.Raise 1, MethodName, "Incorrect TabName found"

        If .Item(0).TargetCell <> "$V$16" Or .Item(1).TargetCell <> "$W$16" _
        Or .Item(2).TargetCell <> "$V$16" Or .Item(3).TargetCell <> "$W$16" _
        Then Err.Raise 1, MethodName, "Incorrect Cell found"
        
        MsgBox "Successfully parsed: " & vbNewLine & Formula & "as" & vbNewLine & vbNewLine & _
            "Path: " & .Item(0).TargetPath, _
            vbOKOnly, MethodName
    End With
    
XT: Exit Sub
    
EH: Select Case MsgBoxAbortRetryIgnore(Err, MethodName)
        Case vbRetry:  Resume
        Case vbIgnore: Resume Next
    End Select
    Resume XT
    Resume
End Sub

Public Sub CellParseLinkTest()
    Const MethodName As String = mModuleName & "CellParseLinkTest"
    Const PathPrefix As String = "S:\can\Finance\actuarial\ASSC\Institutional\Reporting\2016\"
    Const Formula As String = _
        "=SUM('S:\can\Finance\actuarial\ASSC\Institutional\Reporting\2016\M03\Reserves\CRR\Affinity\[CCPE Experience Report_2016 Q1.xls]INTERIM REPORT'!$V$16," & _
        "     'S:\can\Finance\actuarial\ASSC\Institutional\Reporting\2016\M03\Reserves\CRR\Affinity\[CCPE Experience Report_2016 Q1.xls]INTERIM REPORT'!$W$16)" & _
        "-SUM('S:\can\finance\actuarial\ASSC\Institutional\Reporting\2015\M12\Reserves\CRR\Affinity\[CCPE Experience Report_2015 Q4 - Corrected.xls]INTERIM REPORT'!$V$16," & _
        "     'S:\can\finance\actuarial\ASSC\Institutional\Reporting\2015\M12\Reserves\CRR\Affinity\[CCPE Experience Report_2015 Q4 - Corrected.xls]INTERIM REPORT'!$W$16)"
    
    On Error GoTo EH
    Dim ExtLinks As IExternalLinks
    Set ExtLinks = AddInHandle.Parse(DummyCellRef, Formula).Links
    With ExtLinks
        If .Item(0).TargetPath <> "S:\can\Finance\actuarial\ASSC\Institutional\Reporting\2016\M03\Reserves\CRR\Affinity\" _
        Or .Item(1).TargetPath <> "S:\can\Finance\actuarial\ASSC\Institutional\Reporting\2016\M03\Reserves\CRR\Affinity\" _
        Or .Item(2).TargetPath <> "S:\can\finance\actuarial\ASSC\Institutional\Reporting\2015\M12\Reserves\CRR\Affinity\" _
        Or .Item(3).TargetPath <> "S:\can\finance\actuarial\ASSC\Institutional\Reporting\2015\M12\Reserves\CRR\Affinity\" _
        Then Err.Raise 1, MethodName, "Incorrect Path found"
        
        If .Item(0).TargetFile <> "CCPE Experience Report_2016 Q1.xls" _
        Or .Item(1).TargetFile <> "CCPE Experience Report_2016 Q1.xls" _
        Or .Item(2).TargetFile <> "CCPE Experience Report_2015 Q4 - Corrected.xls" _
        Or .Item(3).TargetFile <> "CCPE Experience Report_2015 Q4 - Corrected.xls" _
        Then Err.Raise 1, MethodName, "Incorrect FileName found"

        If .Item(0).TargetTab <> "INTERIM REPORT" _
        Or .Item(1).TargetTab <> "INTERIM REPORT" _
        Or .Item(2).TargetTab <> "INTERIM REPORT" _
        Or .Item(3).TargetTab <> "INTERIM REPORT" _
        Then Err.Raise 1, MethodName, "Incorrect TabName found"

        If .Item(0).TargetCell <> "$V$16" Or .Item(1).TargetCell <> "$W$16" _
        Or .Item(2).TargetCell <> "$V$16" Or .Item(3).TargetCell <> "$W$16" _
        Then Err.Raise 1, MethodName, "Incorrect Cell found"

        MsgBox "Successfully parsed: " & vbNewLine & Formula & vbNewLine & "as" & vbNewLine & vbNewLine & _
            "Path: " & .Item(0).TargetPath, _
            vbOKOnly, MethodName
    End With
    
XT: Exit Sub
    
EH: Select Case MsgBoxAbortRetryIgnore(Err, MethodName)
        Case vbRetry:  Resume
        Case vbIgnore: Resume Next
    End Select
    Resume XT
    Resume
End Sub

Public Sub ArrayNamedRangeTest()
    Const MethodName As String = mModuleName & "ArrayNamedRangeTest", _
          Literal1   As String = """Written Quote Out""", _
          Literal2   As String = """Accepted Quotes""", _
          Literal3   As String = """Rejected"""
    Const Formula As String = _
        "={#N/A,#N/A,FALSE," & Literal1 & _
        ";#N/A,#N/A,FALSE," & Literal2 & _
        ";#N/A,#N/A,FALSE," & Literal3 & "}"
    
    On Error GoTo EH
    Dim Lexer As ILinksLexer: Set Lexer = NewLinksLexer(DummyCellRef, Formula)
    VerifyToken Lexer, EToken_Equals, "="
    VerifyToken Lexer, EToken_OpenBrace, "{"
    VerifyToken Lexer, EToken_Identifier, "#N/A":      VerifyToken Lexer, EToken_Comma, ","
    VerifyToken Lexer, EToken_Identifier, "#N/A":      VerifyToken Lexer, EToken_Comma, ","
    VerifyToken Lexer, EToken_Identifier, "FALSE":     VerifyToken Lexer, EToken_Comma, ","
    VerifyToken Lexer, EToken_StringLiteral, Literal1: VerifyToken Lexer, EToken_SemiColon, ";"
    
    VerifyToken Lexer, EToken_Identifier, "#N/A":      VerifyToken Lexer, EToken_Comma, ","
    VerifyToken Lexer, EToken_Identifier, "#N/A":      VerifyToken Lexer, EToken_Comma, ","
    VerifyToken Lexer, EToken_Identifier, "FALSE":     VerifyToken Lexer, EToken_Comma, ","
    VerifyToken Lexer, EToken_StringLiteral, Literal2: VerifyToken Lexer, EToken_SemiColon, ";"
    
    VerifyToken Lexer, EToken_Identifier, "#N/A":      VerifyToken Lexer, EToken_Comma, ","
    VerifyToken Lexer, EToken_Identifier, "#N/A":      VerifyToken Lexer, EToken_Comma, ","
    VerifyToken Lexer, EToken_Identifier, "FALSE":     VerifyToken Lexer, EToken_Comma, ","
    VerifyToken Lexer, EToken_StringLiteral, Literal3: VerifyToken Lexer, EToken_CloseBrace, "}"
    
    VerifyBraceDepth Lexer, 0
    VerifyParenDepth Lexer, 0
    ScanCheckEOT MethodName, Lexer
    
    MsgBox "Successfully parsed: " & vbNewLine & Formula, vbOKOnly, MethodName
    
XT: Exit Sub
    
EH: Select Case MsgBoxAbortRetryIgnore(Err, MethodName)
        Case vbRetry:  Resume
        Case vbIgnore: Resume Next
    End Select
    Resume XT
    Resume
End Sub

Private Sub ScanCheck(ByVal Test As String, ByVal Lexer As ILinksLexer, _
    ByVal TokenExpected As EToken, ByVal Expected As String _
)
    On Error GoTo EH
    Dim Token As IToken: Set Token = Lexer.Scan()
    If Token.Value <> Token Or Token.Text <> Expected Then _
        Err.Raise 1, Test, _
            vbNewLine & "Expected: '" & Expected & "'" & _
            vbNewLine & "Found:    '" & Token.Text & "'"
XT: Exit Sub
EH: ErrorUtils.ReRaiseError Err, mModuleName & ""
    Resume      ' for debugging only
End Sub

Private Sub ScanCheckEOT(ByVal MethodName As String, ByVal Lexer As ILinksLexer)
    On Error GoTo EH
    Dim Token As IToken: Set Token = Lexer.Scan()
    If Token.Value <> EToken_EOT Then Err.Raise 1, MethodName, "Expected: EOT"
XT: Exit Sub
EH: ErrorUtils.ReRaiseError Err, mModuleName & ""
    Resume      ' for debugging only
End Sub

Private Sub VerifyToken(ByVal Lexer As ILinksLexer, ByVal ExpectedType As EToken, ByVal ExpectedText As String)
    Const MethodName As String = mModuleName & "VerifyNextToken"
    Dim Token As IToken: Set Token = Lexer.Scan()
    If Token.Value <> ExpectedType Or Token.Text <> ExpectedText Then _
            Err.Raise 1, MethodName, _
            vbNewLine & "Expected: '" & ExpectedText & "'" & _
            vbNewLine & "Found:    '" & Token.Text & "'"
End Sub
Private Function VerifyParenDepth(ByVal Lexer As ILinksLexer, ByVal ExpectedDepth As Long)
    Const MethodName As String = mModuleName & "VerifyParenDepthExpected"
    If Lexer.ParenDepth <> ExpectedDepth Then _
        Err.Raise 1, MethodName, "Paren depth = " & Lexer.ParenDepth & "; expected " & ExpectedDepth
End Function

Private Function VerifyBraceDepth(ByVal Lexer As ILinksLexer, ByVal ExpectedDepth As Long)
    Const MethodName As String = mModuleName & "VerifyParenDepthExpected"
    If Lexer.ParenDepth <> ExpectedDepth Then _
        Err.Raise 1, MethodName, "Brace depth = " & Lexer.ParenDepth & "; expected " & ExpectedDepth
End Function
