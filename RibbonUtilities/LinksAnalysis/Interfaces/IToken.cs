////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////

using System;
using System.Runtime.InteropServices;

namespace PGSolutions.RibbonUtilities.LinksAnalysis.Interfaces {

    /// <summary>TODO</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IToken)]
    public interface IToken {
        int     Start  { get; }
        string  Text   { get; }
        EToken  Value  { get; }
    }
}
