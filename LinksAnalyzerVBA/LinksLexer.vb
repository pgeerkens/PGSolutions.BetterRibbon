Public Class LinksLexer
#Const ShowError = 0

    Private Const mModuleName As String = "Lexer."

    Private mIsInitialized As Boolean
    Private mTextIn As String
    Private mCharPos As Long
    Private mCurrentChar As String

    Private mWordOperators As Collection

    Private mBraceDepth As Long
    Private mParenDepth As Long

    Public Property Get Text() As String
    Text = mTextIn
End Property
    Public Function LoadText(ByVal TextIn As String) As LinksLexer
        Const MethodName As String = mModuleName & "Path"

        On Error GoTo EH
        mTextIn = TextIn
        mCharPos = 0
        Advance()
            
    Set LoadText = Me
    mIsInitialized = True
XT:     Exit Function

EH:     ReRaiseError Err, MethodName
End Function

    Friend Function Scan(ByRef TokenText As String) As Token
        Const MethodName As String = mModuleName & "Scan"
        ErrorUtils.CheckIsInitialized mIsInitialized, MethodName

    On Error GoTo EH
        Do
            If IsEOT Then Scan = EOT : GoTo XT
            Select Case mCurrentChar
                Case vbCr, vbLf, vbTab, " " : Advance()
                Case Else : GoTo NoMoreWhiteSpace
            End Select
        Loop

NoMoreWhiteSpace:
        Dim TokenStart As Long : TokenStart = mCharPos
        Select Case VBA.UCase(mCurrentChar)
            Case "!" : Advance() :
                Scan = Bang
            Case "=" : Advance() :
                Scan = Equals()
            Case "," : Advance() :
                Scan = Comma
            Case ";" : Advance() :
                Scan = SemiColon
            Case "+", "-", "%" : Advance() :
                Scan = Unop
            Case "*", "/", "&", "^", "<", ">" : Advance() :
                Scan = BinOp
            Case "<", ">" : Advance() :
                If mCurrentChar = "=" Then Advance() :
                Scan = BinOp
            Case "(" : Advance() :
                Scan = OpenParen : mParenDepth = mParenDepth + 1
            Case ")" : Advance() :
                Scan = CloseParen : mParenDepth = mParenDepth - 1
            Case "{" : Advance() :
                Scan = OpenBrace : mBraceDepth = mBraceDepth + 1
            Case "}" : Advance() :
                Scan = CloseBrace : mBraceDepth = mBraceDepth - 1
            Case "A" To "Z", "_", "$" : Scan = ScanIdent()
            Case "#" : Scan = ScanErrorIdent()
            Case "0" To "9" : Scan = ScanNumber()
            Case "'" : Scan = ScanExternRef()
            Case """" : Scan = ScanStringLiteral()
            Case Else : Advance() :
                Scan = ScanError
        End Select

        TokenText = VBA.mID$(mTextIn, TokenStart, mCharPos - TokenStart)

        If Scan = Identifier Then If IsWordOperator(TokenText) Then Scan = BinOp
XT:     Exit Function

EH:     ReRaiseError Err, MethodName
End Function

    Friend Sub VerifyNextToken(ByVal ExpectedType As Token, ByVal ExpectedText As String)
        Const MethodName As String = mModuleName & "VerifyNextToken"
        Dim TokenText As String
        If Scan(TokenText) <> ExpectedType Or TokenText <> ExpectedText Then _
            Err.Raise 1, MethodName, "Expected: '" & ExpectedText & "'"
End Sub

    Public Function VerifyParenDepth(ByVal ExpectedDepth As Long)
        Const MethodName As String = mModuleName & "VerifyParenDepthExpected"
        If mParenDepth <> ExpectedDepth Then _
        Err.Raise 1, MethodName, "Paren depth = " & mParenDepth & "; expected " & ExpectedDepth
End Function

    Public Function VerifyBraceDepth(ByVal ExpectedDepth As Long)
        Const MethodName As String = mModuleName & "VerifyParenDepthExpected"
        If mParenDepth <> ExpectedDepth Then _
        Err.Raise 1, MethodName, "Brace depth = " & mBraceDepth & "; expected " & ExpectedDepth
End Function

    Private Function IsWordOperator(ByVal Text As String) As Boolean
        Const MethodName As String = mModuleName & "Scan"

        On Error GoTo EH
        IsWordOperator = False
        IsWordOperator = mWordOperators.Item(Text) <> ""
XT:     Exit Function

EH:     If Err.Number = 5 Then Resume Next  ' Subscript out of range => Not Found
        ReRaiseError Err, MethodName
End Function

    Friend Function ScanExternRef() As Token
        Const MethodName As String = mModuleName & "ScanExternRef"
        ErrorUtils.CheckIsInitialized mIsInitialized, MethodName

    On Error GoTo EH
        ScanExternRef = ScanError
        Do
            If Advance() Then GoTo XT
            If mCurrentChar = "'" Then
                If Advance() Then GoTo XT
                If mCurrentChar <> "'" Then GoTo ExitLoop
            End If
        Loop
ExitLoop:
        ScanExternRef = ExternRef

XT:     Exit Function

EH:     ReRaiseError Err, MethodName
End Function

    Friend Function ScanStringLiteral() As Token
        Const MethodName As String = mModuleName & "ScanStringLiteral"
        ErrorUtils.CheckIsInitialized mIsInitialized, MethodName

    On Error GoTo EH
        ScanStringLiteral = StringLiteral
        Do While Not Advance()

            If mCurrentChar = """" Then
                If Advance() Then GoTo XT
                If mCurrentChar <> """" Then GoTo XT
                If Advance() Then GoTo XT
            End If
        Loop
        ScanStringLiteral = ScanError

XT:     Exit Function

EH:     ReRaiseError Err, MethodName
End Function

    Friend Function ScanIdent() As Token
        Const MethodName As String = mModuleName & "ScanIdent"
        ErrorUtils.CheckIsInitialized mIsInitialized, MethodName

    On Error GoTo EH
        ScanIdent = ScanError
        Do
            Select Case mCurrentChar
                Case "A" To "Z", "_", "$", ":", "0" To "9" ' NO-OP
                Case Else : Exit Do
            End Select
        Loop Until Advance()

        ScanIdent = Identifier
XT:     Exit Function

EH:     ReRaiseError Err, MethodName
End Function

    Friend Function ScanErrorIdent() As Token
        Const MethodName As String = mModuleName & "ScanErrorIdent"
        ErrorUtils.CheckIsInitialized mIsInitialized, MethodName

    On Error GoTo EH
        Advance()
        If ScanIdent() <> Identifier Then GoTo ScanError

        Select Case mCurrentChar
#If ShowError = 1 Then
    Case "!":       Advance
#Else
            Case "!", "?" : Advance()
#End If
                ScanErrorIdent = Identifier
            Case "/" : Advance()
                If ScanIdent() <> Identifier Then GoTo ScanError
                ScanErrorIdent = Identifier
            Case Else : GoTo ScanError
        End Select

XT:     Exit Function

ScanError:
        ScanErrorIdent = ScanError
        GoTo XT

EH:     ReRaiseError Err, MethodName
End Function

    Friend Function ScanNumber() As Token
        Const MethodName As String = mModuleName & "ScanNumber"
        ErrorUtils.CheckIsInitialized mIsInitialized, MethodName

    On Error GoTo EH
        ScanNumber = ScanError
        Dim ParsingFraction As Boolean
        Do
            Select Case mCurrentChar
                Case "0" To "9"                ' NO-OP
                Case "."
                    If ParsingFraction Then ScanNumber = ScanError
                    ParsingFraction = True
                Case Else : Exit Do
            End Select
        Loop Until Advance()

        ScanNumber = Number
XT:     Exit Function

EH:     ReRaiseError Err, MethodName
End Function

    Public Property Get CharPos() As Long
    ErrorUtils.CheckIsInitialized mIsInitialized, mModuleName & "CharPos"
    CharPos = mCharPos
End Property

    ''' <summary>Returns true exactly when the last token on the string has been scanned.</summary>
    Public Property Get IsEOT() As Boolean
    IsEOT = mCharPos > Len(mTextIn)
End Property
    Public Property Get This() As LinksLexer
    ErrorUtils.CheckIsInitialized mIsInitialized, mModuleName & "CharPos"
    Set This = Me
End Property

    Public Function RaiseError(
    ByVal CellRef As InternalCellRef,
    ByVal ExpectedText As String
) As ParseError
        With New ParseError
        Set RaiseError = .Initialize(CellRef, mTextIn, CharPos, _
                ExpectedText & " at position " & CharPos)
    End With
    End Function

    ''' <summary>Advances current character by one space, and returns the value of IsEOT.</summary>
    ''' <returns>Returns true exactly when the last token on the string has been scanned.</returns>
    Private Function Advance() As Boolean
        If Not IsEOT Then
            mCharPos = mCharPos + 1
            If Not IsEOT Then mCurrentChar = VBA.UCase$(VBA.mID$(mTextIn, mCharPos, 1))
        End If
        Advance = IsEOT
    End Function

    Private Sub Class_Initialize()
        mIsInitialized = False
    Set mWordOperators = New Collection
    mWordOperators.Add "AND", "AND"
    mWordOperators.Add "OR", "OR"
End Sub

End Class
