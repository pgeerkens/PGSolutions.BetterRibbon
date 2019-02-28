////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace PGSolutions.RibbonDispatcher.ViewModels {
    /// <summary>The interface for the Ribbon ViewModelFactory.</summary>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design","CA1040:AvoidEmptyInterfaces")]
    [Description("The factory interface for the Ribbon ModelFactory.")]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [ComVisible(true), Guid(Guids.IViewModelFactory)]
    public interface IViewModelFactory {
    }
}
