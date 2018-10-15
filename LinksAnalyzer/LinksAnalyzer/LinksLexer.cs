////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

using PGSolutions.LinksAnalyzer.Interfaces;

namespace PGSolutions.LinksAnalyzer {
    /// <summary>TODO</summary>
    [Serializable]
    [CLSCompliant(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ILinksLexer))]
    public class LinksLexer : ILinksLexer {
        public LinksLexer(ISourceCellRef cellRef, string formula) {
            CellRef      = cellRef;
            Formula      = formula + " ";
            ParenDepth   = 0;
            BraceDepth   = 0;
            CharPosition = 0;
        }

        public static IReadOnlyList<string> WordOperators = new List<string> { "AND", "OR" };

        public  ISourceCellRef   CellRef      { get; }
        public  string           Formula      { get; }
        public  int              CharPosition { get; private set; }
        public  int              ParenDepth   { get; private set; }
        public  int              BraceDepth   { get; private set; }

        private char   CurrentCharacter {  get; set; }
        private bool   IsEOT            => CharPosition >= Formula.Length;

        string GetText(int start) => Formula.Substring(start-1, CharPosition + 1 - start);

        public IToken Scan() {
            while ( Advancable() ) {
                if ( ! CurrentCharacter.IsWhiteSpace() ) {
                    var start = CharPosition;
                    switch (CurrentCharacter) {
                        case '!': return EToken.Bang.Set(start, GetText(start));
                        case '=': return EToken.Equals.Set(start, GetText(start));
                        case ',': return EToken.Comma.Set(start, GetText(start));
                        case ';': return EToken.SemiColon.Set(start, GetText(start));
                        case '+':
                        case '-':
                        case '%': return EToken.Unop.Set(start, GetText(start));
                        case '*':
                        case '/':
                        case '&':
                        case '^': return EToken.BinOp.Set(start, GetText(start));
                        case '<':
                        case '>': if ( IsEOT ) { break; }
                                  if (NextCharacterIs('=')) { CharPosition++; }
                                  return EToken.BinOp.Set(start, GetText(start));
                        case '(': ParenDepth++; return EToken.OpenParen.Set(start, GetText(start));
                        case ')': ParenDepth--; return EToken.CloseParen.Set(start, GetText(start));
                        case '{': BraceDepth++; return EToken.OpenBrace.Set(start, GetText(start));
                        case '}': BraceDepth--; return EToken.CloseBrace.Set(start, GetText(start));
                        case '#': return ScanErrorIdent(start);
                        case '"': return ScanStringLiteral(start);
                        case '\'': return ScanExternalRef(start);
                        default:
                            if (CurrentCharacter.IsNumeric()) { return ScanNumber(start); }
                            if ( CurrentCharacter.IsAlpha() ) { return ScanIdentifier(start); }
                            return EToken.ScanError.Set(start, this);
                    }
                }
            }
            return IsEOT ? EToken.EOT.Set(CharPosition, this)
                         : EToken.ScanError.Set(CharPosition, this);
        }

        private IToken ScanExternalRef(int start) {
            while ( Advancable() ) {
                 if (CurrentCharacter != '\'') {
                    continue;
                } else if (NextCharacterIs('!')) {
                    return EToken.ExternRef.Set(start, GetText(start));
                } else {
                    return EToken.ScanError.Set(start, this);
                }
            }
            return EToken.ScanError.Set(start, this);
        }

        private IToken ScanStringLiteral(int start) {
            while ( Advancable() ) {
                if (CurrentCharacter == '"') {
                    if (! Advancable()  ||  CurrentCharacter.IsWhiteSpace()) {
                        return EToken.StringLiteral.Set(start, this);
                    } else if (CurrentCharacter == '"') {
                        continue;
                    } else {
                        return EToken.StringLiteral.Set(start, Formula.Substring(start-1, CharPosition-- - start));
                    }
                }
            }
            return EToken.ScanError.Set(start, this);
        }

        private IToken ScanIdentifier(int start)  => ScanIdent(start, c => c==':');

        private IToken ScanErrorIdent(int start) => ScanIdent(start, c => c==':' || c=='/');

         private IToken ScanIdent(int start, Func<char,bool> extraChars) {
            while ( Advancable()) {
                if (CurrentCharacter.IsAlphanumeric()  ||  extraChars(CurrentCharacter) ) {
                    continue;
                } else if (CurrentCharacter.IsWhiteSpace()) {
                    return EToken.Identifier.Set(start, this);
                } else {
                    return EToken.Identifier.Set(start, Formula.Substring(start-1, CharPosition-- - start));
                }
            }
            return EToken.Identifier.Set(start, this);
        }

       private IToken ScanNumber(int start) {
            bool ParsingFraction = false;
            while (Advancable()) {
                if (CurrentCharacter.IsNumeric()) {
                    continue;
                } else if (CurrentCharacter.IsWhiteSpace()) {
                    return EToken.Number.Set(start, this);
                } else if (CurrentCharacter != '.') {
                    return EToken.Number.Set(start, Formula.Substring(start-1, CharPosition-- - start));
                } else if (ParsingFraction) {
                    return EToken.ScanError.Set(start, this);
                } else {
                    ParsingFraction = true;
                }
            }
            return EToken.Number.Set(start, this);
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
    }
}
