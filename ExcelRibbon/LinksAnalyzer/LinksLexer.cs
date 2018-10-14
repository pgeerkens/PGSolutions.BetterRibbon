using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelRibbon.LinksAnalyzer {
    public enum Token { 
        ScanError,
        EOT,
        Identifier,
        StringLiteral,
        Number,
        BinOp,
        Unop,
        Equals,
        Comma,
        SemiColon,
        Bang,
        OpenParen,
        CloseParen,
        ExternRef,
        OpenBrace,
        CloseBrace
    }

    public interface IToken {
        Token   Token   { get; }
        string  Text    { get; }
        long    Start   { get;}
    }

    internal static partial class Extensions {
        public static IToken Set(this Token token, int start, LinksLexer lexer)
            => lexer.Set(token, start, lexer.CharPosition);

        public static bool IsNumeric (this char c) => '0' <= c && c <= '9';
        public static bool IsAlpha (this char c) => 'A' <= c && c <= 'Z';
        public static bool IsAlphanumeric(this char c) => c.IsAlpha() || c.IsNumeric();
        public static bool IsWhiteSpace(this char c) => c == ' ' || c == '\n' || c == '\r' || c == '\t';
    }

    public interface ILinksLexer {

    }

    public class LinksLexer {
        public LinksLexer(InternalCellRef cellRef, string textIn) {
            ParenDepth = BraceDepth = 0;
            CellRef      = cellRef;
            TextIn       = textIn;
            CharPosition = 0;
        }

        private static IReadOnlyList<string> WordOperators = new List<string> { "AND", "OR" };

        private char CurrentCharacter;

        private InternalCellRef CellRef { get; set; }
        public  string TextIn       { get; }
        public  int    CharPosition { get; private set; }
        public  bool   IsEOT        => CharPosition >= TextIn.Length;
        private int    ParenDepth   { get; set; }
        private int    BraceDepth   { get; set; }
        private string TokenText    { get; set; }

        private string  Formula => TextIn;

        private ParseError ErrorHere(string source, string condition) =>
            new ParseError(CellRef, Formula, CharPosition, $"{source}: {condition}");

        internal IToken Scan() {
            while ( Advancable() ) {
                if ( ! IsWhiteSpace ) {
                    var start = CharPosition;
                    switch (CurrentCharacter) {
                        case '!': return Token.Bang.Set(start, this);
                        case '=': return Token.Equals.Set(start, this);
                        case ',': return Token.Comma.Set(start, this);
                        case ';': return Token.SemiColon.Set(start, this);
                        case '+':
                        case '-':
                        case '%': return Token.Unop.Set(start, this);
                        case '*':
                        case '/':
                        case '&':
                        case '^': return Token.BinOp.Set(start, this);
                        case '<':
                        case '>': if ( IsEOT ) { break; }
                                  if (NextCharacterIs('=')) { CharPosition++; }
                                  return Token.BinOp.Set(start, this);
                        case '(': ParenDepth++; return Token.OpenParen.Set(start, this);
                        case ')': ParenDepth--; return Token.CloseParen.Set(start, this);
                        case '{': BraceDepth++; return Token.OpenBrace.Set(start, this);
                        case '}': BraceDepth--; return Token.CloseBrace.Set(start, this);
                        case '#': return ScanErrorIdent(start);
                        case '"': return ScanStringLiteral(start);
                        case '\'': return ScanExternalRef(start);
                        default:
                            if (CurrentCharacter.IsNumeric()) return ScanNumber(start);
                            if ( CurrentCharacter.IsAlpha() ) return ScanIdent(start);
                            return Set(Token.ScanError, start, CharPosition);
                    }
                }
            }
            return Token.EOT.Set(CharPosition, this);
        }

        private IToken ScanExternalRef(int start) {
            while ( ! Advancable() ) {
                if (CurrentCharacter == '\'') {
                    if (Advancable() ) { 
                        return Token.ScanError.Set(start, this);
                    } else  if (CurrentCharacter != '\'') {
                        return Token.ExternRef.Set(start, this);
                    }
                }
            }
            return Token.ExternRef.Set(start, this);
        }


        private IToken ScanStringLiteral(int start) {
            while ( Advancable() ) {
                if (CurrentCharacter == '"') {
                    if (! Advancable()  ||  IsWhiteSpace) {
                        return Token.StringLiteral.Set(start, this);
                    } else if (CurrentCharacter == '"') {
                        continue;
                    } else {
                        return Token.StringLiteral.Set(start, this);
                    }
                }
            }
            return Set(Token.ScanError, start, CharPosition);
        }

        internal IToken ScanIdent(int start) {
            while ( Advancable()) {
                if (IsWhiteSpace) { return Token.Identifier.Set(start, this); }
                if (CurrentCharacter.IsAlphanumeric() 
                ||  CurrentCharacter == '_'
                ||  CurrentCharacter == ':') {
                    continue;
                } else {
                    return Token.Identifier.Set(start, this);
                }
            }
            return Token.Identifier.Set(start, this);
        }

        internal IToken ScanErrorIdent(int start) => ScanIdent(start);

        internal IToken ScanNumber(int start) {
            bool ParsingFraction = false;
            while (Advancable()) {
                if (CurrentCharacter.IsNumeric()) {
                    continue;
                } else if (CurrentCharacter != '.') {
                    return Token.Number.Set(start, this);
                } else if (ParsingFraction) {
                    return Token.ScanError.Set(start, this);
                } else {
                    ParsingFraction = true;
                }
            }
            return Token.Number.Set(start, this);
        }

        internal IToken Set(Token token, int start, int end) =>
            new TokenX(token, start, TextIn.Substring(start, end - start));

        /// <summary>If not EOT scans forward snd returns true if that charaacter is white space..</summary>
        /// <returns></returns>
        private bool Advancable() {
            if ( IsEOT ) { 
                return false;
            } else {
                CurrentCharacter = TextIn[CharPosition++];
                return true;
            }
        }

        private bool NextCharacterIs(char c) => ! IsEOT && TextIn[CharPosition + 1] == c;

        private bool IsWhiteSpace => 
            CurrentCharacter == '\r' || CurrentCharacter == '\n' ||
            CurrentCharacter == '\t' || CurrentCharacter == ' ';

        internal class TokenX : IToken {
            public TokenX(Token token, int start, string text) {
                Token = token;
                Start = start;
                Text  = text;
            }

            public Token    Token   { get; }
            public long     Start   { get;}
            public string   Text    { get; }
        }

        private bool IsWordOperator(string text) => WordOperators.FirstOrDefault(s => s == text) != null;

        internal ParseError VerifyNextToken(Token ExpectedType, string ExpectedText) {
            if (Scan(TokenText).Token == ExpectedType && TokenText == ExpectedText) return null;
            return ErrorHere("VerifyNextToken", $"Expected: '{ExpectedText}'");
        }

        public ParseError VerifyBraceDepth(int expected) {
            if (BraceDepth == expected) return null;
            return ErrorHere("VerifyBraceDepth", $" depth = {BraceDepth}; expected {expected}.");
        }
        public ParseError VerifyParenDepth(int expected) {
            if (ParenDepth == expected) return null;
            return ErrorHere("VerifyParenDepth", $" depth = {ParenDepth}; expected {expected}.");
        }

        public ParseError RaiseError(InternalCellRef cellRef, string expectedText) =>
            new ParseError(cellRef, TextIn, CharPosition, $"{expectedText} at position {CharPosition}");
    }
}
