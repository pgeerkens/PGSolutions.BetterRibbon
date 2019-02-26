////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary>The interface for the Ribbon ViewModelFactory.</summary>
    [Description("The factory interface for the Ribbon ModelFactory.")]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [ComVisible(true), Guid(Guids.IViewModelFactory)]
    public interface IViewModelFactory {
        IResourceLoader ResourceManager { get; }
    }
}
