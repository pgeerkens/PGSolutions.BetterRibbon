﻿////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

using System.Runtime.InteropServices;

namespace PGSolutions.RibbonUtilities.LinksAnalysis.Interfaces {
    /// <summary>TODO</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IExternalNamedRef)]
    public interface ICellRef {
        string Formula      { get; }
        string TargetPath   { get; }
        string TargetFile   { get; }
        string TargetTab    { get; }
        string TargetCell   { get; }

        bool   IsNamedRange { get; }
        string SourcePath   { get; }
        string SourceFile   { get; }
        string SourceTab    { get; }
        string SourceCell   { get; }
    }

    /// <summary>TODO</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.ILinksLexer)]
    public interface ILinksLexer {
        ISourceCellRef CellRef      { get; }
        int            CharPosition { get; }
        string         Formula      { get; }
        int            ParenDepth   { get; }
        int            BraceDepth   { get; }

        IToken Scan();
    }
}
