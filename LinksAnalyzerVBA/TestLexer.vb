Public Class TestLexer

    Private Const mModuleName As String = "TestLexer."

    Public Sub TestAll()
        SimpleOperatorTest()
        SimpleConcatTest()
        SimpleParensTest()
        StringLiteralTest()
        ComplexRefTest()
        SimpleParseLinkTest()
        ComplexParseLinkTest()
        CellParseLinkTest()
        ArrayNamedRangeTest()
    End Sub

    Private Sub SimpleOperatorTest()
        Const MethodName As String = mModuleName & "SimpleOperatorTest"

        On Error GoTo EH
        Dim TokenText As String
        With New LinksLexer
            Const TestText As String = "4 + 5"
            .LoadText TestText
        ScanCheck MethodName, .This, Number, "4"
        ScanCheck MethodName, .This, Unop, "+"
        ScanCheck MethodName, .This, Number, "5"
        If .Scan(TokenText) <> EOT Then Err.Raise 1, MethodName, "Expected: EOT"
    End With
        MsgBox "Successfully scanned: " & vbNewLine & TestText, vbOKOnly, MethodName
XT:     Exit Sub

EH:     Select Case MsgBoxAbortRetryIgnore(Err, MethodName)
            Case vbRetry : Resume
            Case vbIgnore : Resume Next
        End Select
        Resume XT
        Resume
    End Sub

    Private Sub SimpleConcatTest()
        Const MethodName As String = mModuleName & "SimpleConcatTest"

        On Error GoTo EH
        Dim TokenText As String
        With New LinksLexer
            Const TestText As String = "B4&"" YTD"""
            .LoadText TestText
        If .Scan(TokenText) <> Identifier Or TokenText <> "B4" Then _
                                    Err.Raise 1, MethodName, "Expected: '('"
        If .Scan(TokenText) <> BinOp Or TokenText <> "&" Then _
                                    Err.Raise 1, MethodName, "Expected: '" '"
            If .Scan(TokenText) <> StringLiteral Or TokenText <> """ YTD""" Then _
                                    Err.Raise 1, MethodName, "Expected: '"" YTD""'"

        If .Scan(TokenText) <> EOT Then _
                                    Err.Raise 1, MethodName, "Expected: EOT"
    End With
        MsgBox "Successfully scanned: " & vbNewLine & TestText, vbOKOnly, MethodName
XT:     Exit Sub

EH:     Select Case MsgBoxAbortRetryIgnore(Err, MethodName)
            Case vbRetry : Resume
            Case vbIgnore : Resume Next
        End Select
        Resume XT
        Resume
    End Sub

    Private Sub SimpleParensTest()
        Const MethodName As String = mModuleName & "SimpleParensTest"

        On Error GoTo EH
        Dim TokenText As String
        With New LinksLexer
            Const TestText As String = "(4+5)"
            .LoadText TestText
        If .Scan(TokenText) <> OpenParen Or TokenText <> "(" Then _
                                    Err.Raise 1, MethodName, "Expected: '('"
        If .Scan(TokenText) <> Number Or TokenText <> "4" Then _
                                    Err.Raise 1, MethodName, "Expected: '4'"
        If .Scan(TokenText) <> Unop Or TokenText <> "+" Then _
                                    Err.Raise 1, MethodName, "Expected: '+'"
        If .Scan(TokenText) <> Number Or TokenText <> "5" Then _
                                    Err.Raise 1, MethodName, "Expected: '5'"
        If .Scan(TokenText) <> CloseParen Or TokenText <> ")" Then _
                                    Err.Raise 1, MethodName, "Expected: ')'"

        If .Scan(TokenText) <> EOT Then _
                                    Err.Raise 1, MethodName, "Expected: EOT"
    End With
        MsgBox "Successfully scanned: " & vbNewLine & TestText, vbOKOnly, MethodName
XT:     Exit Sub

EH:     Select Case MsgBoxAbortRetryIgnore(Err, MethodName)
            Case vbRetry : Resume
            Case vbIgnore : Resume Next
        End Select
        Resume XT
        Resume
    End Sub

    Private Sub StringLiteralTest()
        Const MethodName As String = mModuleName & "StringLiteralTest"

        On Error GoTo EH
        Dim TokenText As String
        With New LinksLexer
            Const TestText As String = "=MID(C2, FIND("" '"", C2, 1)+1,  FIND(""]"", C2, 1)-FIND(""'"", C2, 1))"
            .LoadText TestText
        ScanCheck MethodName, .This, Token.Equals, "="
        ScanCheck MethodName, .This, Token.Identifier, "MID"
        ScanCheck MethodName, .This, Token.OpenParen, "("
        ScanCheck MethodName, .This, Token.Identifier, "C2"
        ScanCheck MethodName, .This, Token.Comma, ","
        ScanCheck MethodName, .This, Token.Identifier, "FIND"
        ScanCheck MethodName, .This, Token.OpenParen, "("
        ScanCheck MethodName, .This, Token.StringLiteral, """ '"""
        ScanCheck MethodName, .This, Token.Comma, ","
        ScanCheck MethodName, .This, Token.Identifier, "C2"

        ScanCheck MethodName, .This, Token.Comma, ","
        ScanCheck MethodName, .This, Token.Number, "1"
        ScanCheck MethodName, .This, Token.CloseParen, ")"
        ScanCheck MethodName, .This, Token.Unop, "+"
        ScanCheck MethodName, .This, Token.Number, "1"
        ScanCheck MethodName, .This, Token.Comma, ","
        ScanCheck MethodName, .This, Token.Identifier, "FIND"
        ScanCheck MethodName, .This, Token.OpenParen, "("
        ScanCheck MethodName, .This, Token.StringLiteral, """]"""
        ScanCheck MethodName, .This, Token.Comma, ","

        ScanCheck MethodName, .This, Token.Identifier, "C2"
        ScanCheck MethodName, .This, Token.Comma, ","
        ScanCheck MethodName, .This, Token.Number, "1"
        ScanCheck MethodName, .This, Token.CloseParen, ")"
        ScanCheck MethodName, .This, Token.Unop, "-"
        ScanCheck MethodName, .This, Token.Identifier, "FIND"
        ScanCheck MethodName, .This, Token.OpenParen, "("
        ScanCheck MethodName, .This, Token.StringLiteral, """'"""
        ScanCheck MethodName, .This, Token.Comma, ","
        ScanCheck MethodName, .This, Token.Identifier, "C2"

        ScanCheck MethodName, .This, Token.Comma, ","
        ScanCheck MethodName, .This, Token.Number, "1"
        ScanCheck MethodName, .This, Token.CloseParen, ")"
        ScanCheck MethodName, .This, Token.CloseParen, ")"

        If .Scan(TokenText) <> EOT Then Err.Raise 1, MethodName, "Expected: EOT"
    End With
        MsgBox "Successfully scanned: " & vbNewLine & TestText, vbOKOnly, MethodName
XT:     Exit Sub

EH:     Select Case MsgBoxAbortRetryIgnore(Err, MethodName)
            Case vbRetry : Resume
            Case vbIgnore : Resume Next
        End Select
        Resume XT
        Resume
    End Sub

    Private Sub ComplexRefTest()
        Const MethodName As String = mModuleName & "ComplexRefTest"

        On Error GoTo EH
        Dim TokenText As String
        With New LinksLexer
            Const TestText As String =
            "=VLOOKUP(A18,'S:\can\Affinity\actuar\SPONSOR\VALN\Income Statement Mapping\[IS Mapping.xlsx]IS_line names'!$A$6:$B$400,2,FALSE)"
            .LoadText TestText
        If .Scan(TokenText) <> Equals() Or TokenText <> "=" Then Err.Raise 1, MethodName, "Expected: ="
        If .Scan(TokenText) <> Identifier Or TokenText <> "VLOOKUP" Then Err.Raise 1, MethodName, "Expected: VLOOKUP"
        If .Scan(TokenText) <> OpenParen Or TokenText <> "(" Then Err.Raise 1, MethodName, "Expected: ("
        If .Scan(TokenText) <> Identifier Or TokenText <> "A18" Then Err.Raise 1, MethodName, "Expected: A18"
        If .Scan(TokenText) <> Comma Or TokenText <> "," Then Err.Raise 1, MethodName, "Expected: ,"

        If .Scan(TokenText) <> ExternRef _
        Or TokenText <> "'S:\can\Affinity\actuar\SPONSOR\VALN\Income Statement Mapping\[IS Mapping.xlsx]IS_line names'" Then _
                Err.Raise 1, MethodName,
                "Expected: 'S:\can\Affinity\actuar\SPONSOR\VALN\Income Statement Mapping\[IS Mapping.xlsx]IS_line names'"
        If .Scan(TokenText) <> Bang Or TokenText <> "!" Then Err.Raise 1, MethodName, "Expected: !"
        If .Scan(TokenText) <> Identifier Or TokenText <> "$A$6:$B$400" Then Err.Raise 1, MethodName, "Expected: $A$6:$B$400"
        If .Scan(TokenText) <> Comma Or TokenText <> "," Then Err.Raise 1, MethodName, "Expected: ,"

        If .Scan(TokenText) <> Number Or TokenText <> "2" Then Err.Raise 1, MethodName, "Expected: 2"
        If .Scan(TokenText) <> Comma Or TokenText <> "," Then Err.Raise 1, MethodName, "Expected: ,"
        If .Scan(TokenText) <> Identifier Or TokenText <> "FALSE" Then Err.Raise 1, MethodName, "Expected: False"
        If .Scan(TokenText) <> CloseParen Or TokenText <> ")" Then Err.Raise 1, MethodName, "Expected: )"

        If .Scan(TokenText) <> EOT Then Err.Raise 1, MethodName, "Expected: EOT"
    End With
        MsgBox "Successfully scanned+: " & vbNewLine & TestText, vbOKOnly, MethodName
XT:     Exit Sub

EH:     Select Case MsgBoxAbortRetryIgnore(Err, MethodName)
            Case vbRetry : Resume
            Case vbIgnore : Resume Next
        End Select
        Resume XT
        Resume
    End Sub

    Private Sub SimpleParseLinkTest()
        Const MethodName As String = mModuleName & "SimpleParseLinkTest"

        On Error GoTo EH
        With New ExternalLinks
            Const TestText As String =
            "=VLOOKUP(A18,'S:\can\Affinity\actuar\SPONSOR\VALN\Income Statement Mapping\[IS Mapping.xlsx]IS_line names'!$A$6:$B$400,2,FALSE)"
            Dim Location As InternalCellRef :  Set Location = DummyLocation()
                
        Dim Lexer As LinksLexer :  Set Lexer = New LinksLexer
        .Parse Lexer.LoadText(TestText), Location


        If .ItemByIndex(1).Path <> "S:\can\Affinity\actuar\SPONSOR\VALN\Income Statement Mapping\" _
        Then Err.Raise 1, MethodName, "Incorrect Path found"

        If .ItemByIndex(1).FileName <> "IS Mapping.xlsx" _
        Then Err.Raise 1, MethodName, "Incorrect FileName found"

        If .ItemByIndex(1).TabName <> "IS_line names" _
        Then Err.Raise 1, MethodName, "Incorrect TabName found"

        If .ItemByIndex(1).Cell <> "A6:B400" _
        Then Err.Raise 1, MethodName, "Incorrect Cell found"

        MsgBox "Successfully parsed: " & vbNewLine & TestText & "as" & vbNewLine & vbNewLine &
            "Path: " & .ItemByIndex(1).Path,
            vbOKOnly, MethodName
    End With

XT:     Exit Sub

EH:     Select Case MsgBoxAbortRetryIgnore(Err, MethodName)
            Case vbRetry : Resume
            Case vbIgnore : Resume Next
        End Select
        Resume XT
        Resume
    End Sub

    Private Sub ComplexParseLinkTest()
        Const MethodName As String = mModuleName & "ComplexParseLinkTest"
        Const PathPrefix As String = "S:\can\Finance\actuarial\ASSC\Institutional\Reporting\2016\"

        On Error GoTo EH
        With New ExternalLinks
            Const TestText As String =
            "=SUM('S:\can\Finance\actuarial\ASSC\Institutional\Reporting\2016\M03\Reserves\CRR\Affinity\[CCPE Experience Report_2016 Q1.xls]INTERIM REPORT'!$V$16," & vbNewLine &
            "     'S:\can\Finance\actuarial\ASSC\Institutional\Reporting\2016\M03\Reserves\CRR\Affinity\[CCPE Experience Report_2016 Q1.xls]INTERIM REPORT'!$W$16)" & vbNewLine &
            "-SUM('S:\can\finance\actuarial\ASSC\Institutional\Reporting\2015\M12\Reserves\CRR\Affinity\[CCPE Experience Report_2015 Q4 - Corrected.xls]INTERIM REPORT'!$V$16," & vbNewLine &
            "     'S:\can\finance\actuarial\ASSC\Institutional\Reporting\2015\M12\Reserves\CRR\Affinity\[CCPE Experience Report_2015 Q4 - Corrected.xls]INTERIM REPORT'!$W$16)" & vbNewLine
            Dim Location As InternalCellRef :  Set Location = DummyLocation()
        
        Dim Lexer As LinksLexer :  Set Lexer = New LinksLexer
        .Parse Lexer.LoadText(TestText), Location

        If .ItemByIndex(1).Path <> "S:\can\Finance\actuarial\ASSC\Institutional\Reporting\2016\M03\Reserves\CRR\Affinity\" _
        Or .ItemByIndex(2).Path <> "S:\can\Finance\actuarial\ASSC\Institutional\Reporting\2016\M03\Reserves\CRR\Affinity\" _
        Or .ItemByIndex(3).Path <> "S:\can\finance\actuarial\ASSC\Institutional\Reporting\2015\M12\Reserves\CRR\Affinity\" _
        Or .ItemByIndex(4).Path <> "S:\can\finance\actuarial\ASSC\Institutional\Reporting\2015\M12\Reserves\CRR\Affinity\" _
        Then Err.Raise 1, MethodName, "Incorrect Path found"

        If .ItemByIndex(1).FileName <> "CCPE Experience Report_2016 Q1.xls" _
        Or .ItemByIndex(2).FileName <> "CCPE Experience Report_2016 Q1.xls" _
        Or .ItemByIndex(3).FileName <> "CCPE Experience Report_2015 Q4 - Corrected.xls" _
        Or .ItemByIndex(4).FileName <> "CCPE Experience Report_2015 Q4 - Corrected.xls" _
        Then Err.Raise 1, MethodName, "Incorrect FileName found"

        If .ItemByIndex(1).TabName <> "INTERIM REPORT" _
        Or .ItemByIndex(2).TabName <> "INTERIM REPORT" _
        Or .ItemByIndex(3).TabName <> "INTERIM REPORT" _
        Or .ItemByIndex(4).TabName <> "INTERIM REPORT" _
        Then Err.Raise 1, MethodName, "Incorrect TabName found"

        If .ItemByIndex(1).Cell <> "V16" Or .ItemByIndex(2).Cell <> "W16" _
        Or .ItemByIndex(3).Cell <> "V16" Or .ItemByIndex(4).Cell <> "W16" _
        Then Err.Raise 1, MethodName, "Incorrect Cell found"

        MsgBox "Successfully parsed: " & vbNewLine & TestText & "as" & vbNewLine & vbNewLine &
            "Path: " & .ItemByIndex(1).Path,
            vbOKOnly, MethodName
    End With

XT:     Exit Sub

EH:     Select Case MsgBoxAbortRetryIgnore(Err, MethodName)
            Case vbRetry : Resume
            Case vbIgnore : Resume Next
        End Select
        Resume XT
        Resume
    End Sub

    Private Sub CellParseLinkTest()
        Const MethodName As String = mModuleName & "CellParseLinkTest"
        Const PathPrefix As String = "S:\can\Finance\actuarial\ASSC\Institutional\Reporting\2016\"

        On Error GoTo EH
        With New ExternalLinks
            Dim TestText As String : TestText =
            "=SUM('S:\can\Finance\actuarial\ASSC\Institutional\Reporting\2016\M03\Reserves\CRR\Affinity\[CCPE Experience Report_2016 Q1.xls]INTERIM REPORT'!$V$16," &
            "     'S:\can\Finance\actuarial\ASSC\Institutional\Reporting\2016\M03\Reserves\CRR\Affinity\[CCPE Experience Report_2016 Q1.xls]INTERIM REPORT'!$W$16)" &
            "-SUM('S:\can\finance\actuarial\ASSC\Institutional\Reporting\2015\M12\Reserves\CRR\Affinity\[CCPE Experience Report_2015 Q4 - Corrected.xls]INTERIM REPORT'!$V$16," &
            "     'S:\can\finance\actuarial\ASSC\Institutional\Reporting\2015\M12\Reserves\CRR\Affinity\[CCPE Experience Report_2015 Q4 - Corrected.xls]INTERIM REPORT'!$W$16)"

            Dim Location As InternalCellRef :  Set Location = DummyLocation()
        
        Dim Lexer As LinksLexer :  Set Lexer = New LinksLexer
        .Parse Lexer.LoadText(TestText), Location

        If .ItemByIndex(1).Path <> "S:\can\Finance\actuarial\ASSC\Institutional\Reporting\2016\M03\Reserves\CRR\Affinity\" _
        Or .ItemByIndex(2).Path <> "S:\can\Finance\actuarial\ASSC\Institutional\Reporting\2016\M03\Reserves\CRR\Affinity\" _
        Or .ItemByIndex(3).Path <> "S:\can\finance\actuarial\ASSC\Institutional\Reporting\2015\M12\Reserves\CRR\Affinity\" _
        Or .ItemByIndex(4).Path <> "S:\can\finance\actuarial\ASSC\Institutional\Reporting\2015\M12\Reserves\CRR\Affinity\" _
        Then Err.Raise 1, MethodName, "Incorrect Path found"

        If .ItemByIndex(1).FileName <> "CCPE Experience Report_2016 Q1.xls" _
        Or .ItemByIndex(2).FileName <> "CCPE Experience Report_2016 Q1.xls" _
        Or .ItemByIndex(3).FileName <> "CCPE Experience Report_2015 Q4 - Corrected.xls" _
        Or .ItemByIndex(4).FileName <> "CCPE Experience Report_2015 Q4 - Corrected.xls" _
        Then Err.Raise 1, MethodName, "Incorrect FileName found"

        If .ItemByIndex(1).TabName <> "INTERIM REPORT" _
        Or .ItemByIndex(2).TabName <> "INTERIM REPORT" _
        Or .ItemByIndex(3).TabName <> "INTERIM REPORT" _
        Or .ItemByIndex(4).TabName <> "INTERIM REPORT" _
        Then Err.Raise 1, MethodName, "Incorrect TabName found"

        If .ItemByIndex(1).Cell <> "V16" Or .ItemByIndex(2).Cell <> "W16" _
        Or .ItemByIndex(3).Cell <> "V16" Or .ItemByIndex(4).Cell <> "W16" _
        Then Err.Raise 1, MethodName, "Incorrect Cell found"

        MsgBox "Successfully parsed: " & vbNewLine & TestText & vbNewLine & "as" & vbNewLine & vbNewLine &
            "Path: " & .ItemByIndex(1).Path,
            vbOKOnly, MethodName
    End With

XT:     Exit Sub

EH:     Select Case MsgBoxAbortRetryIgnore(Err, MethodName)
            Case vbRetry : Resume
            Case vbIgnore : Resume Next
        End Select
        Resume XT
        Resume
    End Sub

    Private Sub ArrayNamedRangeTest()
        Const MethodName As String = mModuleName & "ArrayNamedRangeTest"

        On Error GoTo EH
        Dim TestText As String : TestText =
        "={#N/A,#N/A,FALSE,""Written Quote Out"";#N/A,#N/A,FALSE,""Accepted Quotes"";#N/A,#N/A,FALSE,""Rejected""}"
        Dim Location As InternalCellRef :  Set Location = DummyLocation()
    
    With New LinksLexer
            .LoadText TestText
        Dim TokenText As String
            .VerifyNextToken Equals, "="
        .VerifyNextToken OpenBrace, "{"
        .VerifyNextToken Identifier, "#N/A": .VerifyNextToken Comma, ","
        .VerifyNextToken Identifier, "#N/A": .VerifyNextToken Comma, ","
        .VerifyNextToken Identifier, "FALSE": .VerifyNextToken Comma, ","
        .VerifyNextToken StringLiteral, """Written Quote Out""": .VerifyNextToken SemiColon, ";"

        .VerifyNextToken Identifier, "#N/A": .VerifyNextToken Comma, ","
        .VerifyNextToken Identifier, "#N/A": .VerifyNextToken Comma, ","
        .VerifyNextToken Identifier, "FALSE": .VerifyNextToken Comma, ","
        .VerifyNextToken StringLiteral, """Accepted Quotes""": .VerifyNextToken SemiColon, ";"

        .VerifyNextToken Identifier, "#N/A": .VerifyNextToken Comma, ","
        .VerifyNextToken Identifier, "#N/A": .VerifyNextToken Comma, ","
        .VerifyNextToken Identifier, "FALSE": .VerifyNextToken Comma, ","
        .VerifyNextToken StringLiteral, """Rejected""": .VerifyNextToken CloseBrace, "}"

        .VerifyBraceDepth 0
        .VerifyParenDepth 0
    End With

        MsgBox "Successfully parsed: " & vbNewLine & TestText, vbOKOnly, MethodName

XT:     Exit Sub

EH:     Select Case MsgBoxAbortRetryIgnore(Err, MethodName)
            Case vbRetry : Resume
            Case vbIgnore : Resume Next
        End Select
        Resume XT
        Resume
    End Sub

    Private Property Get DummyLocation() As InternalCellRef
    With New InternalCellRef
    Set DummyLocation = .Initialize(ActiveWorkbook.Path, ActiveWorkbook.FullName, "Data", "$A$1")
    End With
    End Property

    Private Sub ScanCheck(Test As String, Lexer As LinksLexer, Token As Token, Expected As String)
        Dim TokenText As String
        If Lexer.Scan(TokenText) <> Token Or TokenText <> Expected Then _
        Err.Raise 1, Test, "Expected: '" & Expected & "'"
End Sub

End Class
