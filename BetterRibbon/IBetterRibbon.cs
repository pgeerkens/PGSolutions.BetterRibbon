////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonUtilities.LinksAnalysis;

namespace PGSolutions.BetterRibbon {
    /// <summary>THe main interface for VBA to access the Ribbon dispatcher.</summary>
    [ComVisible(true)]
    [CLSCompliant(false)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
  //  [Guid(Guids.IBetterRibbon)]
    public interface IBetterRibbon {
        /// <summary>Returns a new implementation of the <see cref="IModelFactory"/> interface.</summary>
        [Description("Returns a new implementation of the IModelFactory interface.")]
        IModelFactory    NewBetterRibbon();
        /// <summary>Returns a new implementation of the <see cref="ILinksAnalyzer"/> interface.</summary>
        [Description("Returns a new implementation of the ILinksAnalyzer interface.")]
        ILinksAnalyzer NewLinksAnalyzer();
    }
}
