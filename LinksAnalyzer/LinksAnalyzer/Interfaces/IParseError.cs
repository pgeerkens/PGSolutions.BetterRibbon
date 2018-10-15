////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

using System.Runtime.InteropServices;

namespace PGSolutions.LinksAnalyzer.Interfaces {
    /// <summary>TODO</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IParseError)]
    public interface IParseError {
        ISourceCellRef CellRef      {  get; }
        string           Formula      { get; }
        int              CharPosition { get; }
        string           Condition    {  get; }
    }
}
