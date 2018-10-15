////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////

using System;
using System.Runtime.InteropServices;

namespace PGSolutions.LinksAnalyzer.Interfaces {
    /// <summary>TODO</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [Guid(Guids.EToken)]
    public enum EToken { 
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
        CloseBrace,
        ErrorTag
    }

    internal static partial class Extensions {
        /// <summary>TODO</summary>
        /// <param name="token"></param>
        /// <param name="start"></param>
        /// <param name="lexer"></param>
        /// <returns></returns>
        public static IToken Set(this EToken token, int start, ILinksLexer lexer)
            => token.Set(start, lexer.Formula.Substring(start-1, lexer.CharPosition - start));

        /// <summary>TODO</summary>
        /// <param name="token"></param>
        /// <param name="start"></param>
        /// <param name="text"></param>
        /// <returns></returns>
        public static IToken Set(this EToken token, int start, string text)
            => text.IsWordOperator() ? new Token(EToken.BinOp, start, text)
                                     : new Token(token, start, text);
    }
}
