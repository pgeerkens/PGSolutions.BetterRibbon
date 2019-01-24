////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
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
}
