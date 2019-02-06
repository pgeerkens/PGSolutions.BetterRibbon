////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////

using System;
using System.Runtime.InteropServices;

namespace PGSolutions.RibbonUtilities.LinksAnalyzer.Interfaces {
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
        BinaryOperator,
        UnaryOperator,
        Equals,
        Comma,
        Semicolon,
        Bang,
        OpenParen,
        CloseParen,
        ExternRef,
        OpenExternRef,
        OpenBrace,
        CloseBrace,
        ErrorTag
    }

    public static partial class Extensions {
        public static string Name(this EToken token) {
            switch (token) {
                case EToken.ScanError:      return "<ScanError>";
                case EToken.EOT:            return "<EOT>";
                case EToken.Identifier:     return "<Identifier>";
                case EToken.StringLiteral:  return "<StringLiteral>";
                case EToken.Number:         return "<Number>";
                case EToken.BinaryOperator: return "<BinOp>";
                case EToken.UnaryOperator:  return "Unop<>";
                case EToken.Equals:         return "<Equals>";
                case EToken.Comma:          return "<Comma>";
                case EToken.Semicolon:      return "<SemiColon>";
                case EToken.Bang:           return "<Bang>";
                case EToken.OpenParen:      return "<OpenParen>";
                case EToken.CloseParen:     return "<CloseParen>";
                case EToken.ExternRef:      return "<ExternRef>";
                case EToken.OpenExternRef:  return "<OpenExternRef>";
                case EToken.OpenBrace:      return "<OpenBrace>";
                case EToken.CloseBrace:     return "<CloseBrace>";
                case EToken.ErrorTag:       return "<ErrorTag>";
                default:                    return "<Unknown>";
            }
        }

    }
}
