////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////

using System;
using System.Runtime.InteropServices;

using PGSolutions.LinksAnalyzer.Interfaces;

namespace PGSolutions.LinksAnalyzer {
    /// <summary>TODO</summary>
    [Serializable]
    [CLSCompliant(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IToken))]
    public class Token : IToken {
        public Token(EToken token, int start, string text) {
            Value = token;
            Start  = start;
            Text   = text;
        }

        public int      Start  { get; }
        public string   Text   { get; }
        public EToken   Value  { get; }
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
