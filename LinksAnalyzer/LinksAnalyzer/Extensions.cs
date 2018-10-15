////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Linq;
using PGSolutions.LinksAnalyzer.Interfaces;

namespace PGSolutions.LinksAnalyzer {
    internal static partial class Extensions {
        public static IParseError ErrorHere(this ILinksLexer lexer, string source, string condition) =>
            new ParseError(lexer.CellRef, lexer.Formula, lexer.CharPosition, $"{source}: {condition}");

        /// <summary>Returns true IFF current character is between '0' and '9'.</summary>
        public static bool IsNumeric     (this char c) => '0' <= c && c <= '9';

        /// <summary>Returns true IFF current character is between 'A' and 'Z, or Underbar or '$''.</summary>
        /// <remarks>These are the valid initial characters for an Identifier token.</remarks>
        public static bool IsAlpha       (this char c) => 'A' <= c && c <= 'Z' || c == '_' || c == '$';

        /// <summary>Returns true IFF current character is Alpha or Numeric.</summary>
        /// <remarks>These are the valid continuation characters for an Identifier token.</remarks>
        public static bool IsAlphanumeric(this char c) => c.IsAlpha() || c.IsNumeric();

        /// <summary>Returns true IFF current character is a LF, CR, TAB, FF, or SPACE.</summary>
        public static bool IsWhiteSpace  (this char c) => c == '\n' || c == '\r'
                                                       || c == '\t' || c == '\f' || c == ' ';

        public static bool IsWordOperator(this string text) => 
            LinksLexer.WordOperators.FirstOrDefault(s => s == text) != null;

        public static string Name(this IToken token) {
            switch (token.Value) {
                case EToken.ScanError:     return "<ScanError>";
                case EToken.EOT:           return "<EOT>";
                case EToken.Identifier:    return "<Identifier>";
                case EToken.StringLiteral: return "<StringLiteral>";
                case EToken.Number:        return "<Number>";
                case EToken.BinOp:         return "<BinOp>";
                case EToken.Unop:          return "Unop<>";
                case EToken.Equals:        return "<Equals>";
                case EToken.Comma:         return "<Comma>";
                case EToken.SemiColon:     return "<SemiColon>";
                case EToken.Bang:          return "<Bang>";
                case EToken.OpenParen:     return "<OpenParen>";
                case EToken.CloseParen:    return "<CloseParen>";
                case EToken.ExternRef:     return "<ExternRef>";
                case EToken.OpenBrace:     return "<OpenBrace>";
                case EToken.CloseBrace:    return "<CloseBrace>";
                case EToken.ErrorTag:      return "<ErrorTag>";
                default:                   return "<Unknown>";
            }
        }
    }
}
