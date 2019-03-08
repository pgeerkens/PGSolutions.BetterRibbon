////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary>THe main interface for VBA to access the Ribbon dispatcher.</summary>
    [ComVisible(true)]
    [CLSCompliant(false)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.ICustomRibbonComEntry)]
    public interface ICustomRibbonComEntry {
        /// <summary>Returns a new implementation of the <see cref="IModelFactory"/> interface.</summary>
        [DispId( 1),Description("Returns a new implementation of the IModelFactory interface.")]
        IModelFactory  NewBetterRibbon(IResourceLoader loader);

        /// <summary>Creates a Ribbon ViewModel for this workbook and registers it with the Dispatcher.</summary>
        [DispId( 3),Description("Creates a Ribbon ViewModel for this workbook and registers it with the Dispatcher.")]
        void RegisterWorkbook(string workbookName);

        /// <summary>Returns a new implementation of the <see cref="IModelServer"/> interface.</summary>
        [DispId( 4),Description("Returns a new implementation of the IModelServer interface.")]
        IModelServer  NewModelServer(IResourceLoader loader);
    }
}
