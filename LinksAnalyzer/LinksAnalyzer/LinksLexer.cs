////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;

using PGSolutions.LinksAnalyzer.Interfaces;

namespace PGSolutions.LinksAnalyzer {
    /// <summary>TODO</summary>
    [System.Diagnostics.CodeAnalysis.SuppressMessage( "Microsoft.Naming", "CA1710:IdentifiersShouldHaveCorrectSuffix" )]
    [Serializable]
    [CLSCompliant(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ILinksLexer))]
    public class LinksLexer : ILinksLexer, IReadOnlyList<IToken> {
        public LinksLexer(ISourceCellRef cellRef, string formula) {
            CellRef      = cellRef;
            Formula      = formula + " ";
            ParenDepth   = 0;
            BraceDepth   = 0;
            CharPosition = 0;
            Tokens       = new List<IToken>();
        }

        public static readonly IReadOnlyList<string> WordOperators = new List<string> { "AND", "OR" };

        public  ISourceCellRef  CellRef      { get; }
        public  string          Formula      { get; }
        public  int             CharPosition { get; private set; }
        public  int             ParenDepth   { get; private set; }
        public  int             BraceDepth   { get; private set; }
        public  int             Count        => Tokens.Count;
        public IToken           this[int index] => Tokens[index];

        private char            CurrentCharacter {  get; set; }
        private bool            IsEOT            => CharPosition >= Formula.Length;
        private IList<IToken>   Tokens {  get; }

        string GetText(int start) => Formula.Substring(start-1, CharPosition + 1 - start);

        [System.Diagnostics.CodeAnalysis.SuppressMessage( "Microsoft.Maintainability", "CA1502:AvoidExcessiveComplexity" )]
        public IToken Scan() {
            if (CharPosition==1 && CurrentCharacter=='\'') { return Add(EToken.StringLiteral, 1, Formula); }
            while ( Advancable() ) {
                if ( ! CurrentCharacter.IsWhiteSpace() ) {
                    var start = CharPosition;
                    switch (CurrentCharacter) {
                        case '!': return Add(EToken.Bang, start, GetText(start));
                        case '=': return Add(EToken.Equals, start, GetText(start));
                        case ',': return Add(EToken.Comma, start, GetText(start));
                        case ';': return Add(EToken.Semicolon, start, GetText(start));
                        case '+':
                        case '-':
                        case '%': return Add(EToken.UnaryOperator, start, GetText(start));
                        case '*':
                        case '/':
                        case '&':
                        case '^': return Add(EToken.BinaryOperator, start, GetText(start));
                        case '<':
                        case '>': if ( IsEOT ) { break; }
                                  if (NextCharacterIs('=')) { CharPosition++; }
                                  return Add(EToken.BinaryOperator, start, GetText(start));
                        case '(': ParenDepth++; return Add(EToken.OpenParen, start, GetText(start));
                        case ')': ParenDepth--; return Add(EToken.CloseParen, start, GetText(start));
                        case '{': BraceDepth++; return Add(EToken.OpenBrace, start, GetText(start));
                        case '}': BraceDepth--; return Add(EToken.CloseBrace, start, GetText(start));
                        case '#': return ScanErrorIdent(start);
                        case '"': return ScanStringLiteral(start);
                        case '\'': return ScanClosedExternalRef(start);
                        case '[': return ScanOpenExternalRef(start);
                        default:
                            if (CurrentCharacter.IsNumeric()) { return ScanNumber(start); }
                            if ( CurrentCharacter.IsAlpha() ) { return ScanIdentifier(start); }
                            return Add(EToken.ScanError, start, this);
                    }
                }
            }
            return IsEOT ? Add(EToken.EOT, CharPosition, this)
                         : Add(EToken.ScanError, CharPosition, this);
        }

        private IToken Add(EToken token, int start, ILinksLexer lexer) =>
            Add(token, start, lexer.Formula.Substring(start-1, lexer.CharPosition - start));

        private IToken Add(EToken token, int start, string text) {
            var rv = text.IsWordOperator() ? new Token(EToken.BinaryOperator, start, text)
                                           : new Token(token, start, text);
            Tokens.Add(rv);
            return rv;
        }

        private IToken ScanClosedExternalRef(int start) {
            while ( Advancable()  && CurrentCharacter != '\'') { }
            return NextCharacterIs('!') ? Add(EToken.ExternRef, start, GetText(start))
                                        : Add(EToken.ScanError, start, this);
        }

        private IToken ScanOpenExternalRef(int start) {
            while ( Advancable()  &&  CurrentCharacter != ']') { }
            while ( Advancable()  &&  CurrentCharacter != '!') {
                // https://www.accountingweb.com/technology/excel/seven-characters-you-cant-use-in-worksheet-names
                switch (CurrentCharacter) {
                    case '/': case '*':
                    case '[': case ']': 
                    case ':': case '?': 
                    case '\'': return Add(EToken.ScanError, start, this);
                    default: break;
                }
            }
            return CurrentCharacter == '!'
                    ? Add(EToken.OpenExternRef, start, Formula.Substring(start-1, CharPosition-- - start))
                    : Add(EToken.ScanError, start, this);
        }

        private IToken ScanStringLiteral(int start) {
            while ( Advancable() ) {
                if (CurrentCharacter == '"') {
                    if (! Advancable()  ||  CurrentCharacter.IsWhiteSpace()) {
                        return Add(EToken.StringLiteral, start, this);
                    } else if (CurrentCharacter == '"') {
                        continue;
                    } else {
                        return Add(EToken.StringLiteral, start, Formula.Substring(start-1, CharPosition-- - start));
                    }
                }
            }
            return Add(EToken.ScanError, start, this);
        }

        private IToken ScanIdentifier(int start)  => ScanIdent(start, c => c==':');

        private IToken ScanErrorIdent(int start) => ScanIdent(start, c => c==':' || c=='/');

         private IToken ScanIdent(int start, Func<char,bool> extraChars) {
            while ( Advancable()) {
                if (CurrentCharacter.IsAlphanumeric()  ||  extraChars(CurrentCharacter) ) {
                    continue;
                } else if (CurrentCharacter.IsWhiteSpace()) {
                    return Add(EToken.Identifier, start, this);
                } else {
                    return Add(EToken.Identifier, start, Formula.Substring(start-1, CharPosition-- - start));
                }
            }
            return Add(EToken.Identifier, start, this);
        }

       private IToken ScanNumber(int start) {
            bool ParsingFraction = false;
            while (Advancable()) {
                if (CurrentCharacter.IsNumeric()) {
                    continue;
                } else if (CurrentCharacter.IsWhiteSpace()) {
                    return Add(EToken.Number, start, this);
                } else if (CurrentCharacter != '.') {
                    return Add(EToken.Number, start, Formula.Substring(start-1, CharPosition-- - start));
                } else if (ParsingFraction) {
                    return Add(EToken.ScanError, start, this);
                } else {
                    ParsingFraction = true;
                }
            }
            return Add(EToken.Number, start, this);
        }

        /// <summary>If not EOT advances CharPosition and returns true; else returns false.</summary>
        private bool Advancable() {
            if ( IsEOT ) { 
                return false;
            } else {
                CurrentCharacter = Formula[CharPosition++];
                return true;
            }
        }

        /// <summary>Returns true IFF not EOT and current character matches that supplied.</summary>
        private bool NextCharacterIs(char c) => ! IsEOT && Formula[CharPosition] == c;

        public IEnumerator<IToken> GetEnumerator() => Tokens.GetEnumerator();
           IEnumerator IEnumerable.GetEnumerator() => Tokens.GetEnumerator();
    }
}
