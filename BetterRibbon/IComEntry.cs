////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonUtilities.LinksAnalysis;

namespace PGSolutions.BetterRibbon {
    /// <summary>THe main interface for VBA to access the Ribbon dispatcher.</summary>
    [ComVisible(true)]
    [CLSCompliant(false)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
  //  [Guid(Guids.IComEntry)]
    public interface IComEntry {
        /// <summary>Returns a new implementation of the <see cref="IModelFactory"/> interface.</summary>
        [DispId( 1),Description("Returns a new implementation of the IModelFactory interface.")]
        IModelFactory  NewBetterRibbon(IResourceLoader manager);

        /// <summary>Returns a new implementation of the <see cref="ILinksAnalyzer"/> interface.</summary>
        [DispId( 2),Description("Returns a new implementation of the ILinksAnalyzer interface.")]
        ILinksAnalyzer NewLinksAnalyzer();

        /// <summary>.</summary>
        [DispId( 3),Description(".")]
        void RegisterWorkbook(string workbookName);

        /// <summary>Returns a new implementation of the <see cref="IModelServer"/> interface.</summary>
        [DispId( 4),Description("Returns a new implementation of the IModelServer interface.")]
        IModelServer  NewModelServer(IResourceLoader manager);
    }
}
