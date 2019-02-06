////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

using System.Runtime.InteropServices;

namespace PGSolutions.RibbonUtilities.LinksAnalyzer.Interfaces {
    /// <summary>TODO</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IParseError)]
    public interface IParseError {
        ISourceCellRef CellRef        { get; }
        string           Formula      { get; }
        int              CharPosition { get; }
        string           Condition    { get; }
    }
    /// <summary>TODO</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.ITwoDimensionalLookup)]
    public interface ITwoDimensionalLookup {
        string   Item(int row, int col);
        int      RowsCount { get; }
        int      ColsCount { get; }
    }
}
