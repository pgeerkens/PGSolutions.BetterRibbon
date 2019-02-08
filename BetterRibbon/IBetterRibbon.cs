﻿////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.InteropServices;
using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonUtilities.LinksAnalysis;

namespace PGSolutions.BetterRibbon {
    /// <summary>THe main interface for VBA to access the Ribbon dispatcher.</summary>
    [ComVisible(true)]
    [CLSCompliant(false)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
  //  [Guid(Guids.IBetterRibbon)]
    public interface IBetterRibbon {
        /// <summary>.</summary>
        IRibbonDispatcher NewBetterRibbon();
        ILinksAnalyzer    NewLinksAnalyzer();
    }
}
