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
}
