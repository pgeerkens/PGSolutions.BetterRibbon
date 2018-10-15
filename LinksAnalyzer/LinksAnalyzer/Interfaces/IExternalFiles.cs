////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace PGSolutions.LinksAnalyzer.Interfaces {
    /// <summary>TODO</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IExternalFiles)]
    public interface IExternalFiles {
        int Count               { get; }
        string this[int index]  { get; }
    }

    /// <summary>TODO</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.ISourceCellRef)]
    public interface ISourceCellRef {
        bool   IsNamedRange { get;}
        string CellName     { get; }
        string TabName      { get; }
        string FileName     { get; }
        string FullPath     { get; }
    }

    /// <summary>TODO</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IParseErrors)]
    public interface IParseErrors {
        int          Count           { get; }
        IParseError  this[int index] { get; }
    }

    /// <summary>TODO</summary>
    [ComVisible(true)]
    [CLSCompliant(false)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IExternalLinks)]
    public interface IExternalLinks {
        int         Count           { get; }
        ICellRef    this[int index] { get; }

        IEnumerator<ICellRef> GetEnumerator();
    }
}
