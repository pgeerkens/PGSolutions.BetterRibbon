////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace PGSolutions.RibbonUtilities.LinksAnalysis.Interfaces {
    /// <summary>TODO</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.ISourceCellRef)]
    public interface ISourceCellRef {
        bool IsNamedRange { get; }
        string CellName { get; }
        string TabName { get; }
        string FileName { get; }
        string FullPath { get; }
    }

    /// <summary>TODO</summary>
    [SuppressMessage( "Microsoft.Naming", "CA1710:IdentifiersShouldHaveCorrectSuffix" )]
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IExternalFiles)]
    public interface IExternalFiles : IReadOnlyList<string> {
        new int Count { get; }
        string  Item(int index);
    }

    /// <summary>TODO</summary>
    [SuppressMessage( "Microsoft.Naming", "CA1710:IdentifiersShouldHaveCorrectSuffix" )]
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IParseErrors)]
    public interface IParseErrors : IReadOnlyList<IParseError> {
        new int     Count { get; }
        IParseError Item(int index);
    }

    /// <summary>TODO</summary>
    [SuppressMessage("Microsoft.Naming", "CA1710:IdentifiersShouldHaveCorrectSuffix")]
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IExternalLinks)]
    public interface IExternalLinks : IReadOnlyList<ICellRef> {
        new int  Count { get; }
        ICellRef Item(int index);
    }

    /// <summary>TODO</summary>
    [ComVisible(true)]
    [CLSCompliant(false)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.ILinksAnalysis)]
    public interface ILinksAnalysis {
        IParseErrors   Errors { get; }
        IExternalFiles Files  { get; }
        IExternalLinks Links  { get; }
    }
}
