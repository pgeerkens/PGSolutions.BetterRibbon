////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary>The interface for the Ribbon ViewModelFactory.</summary>
    [Description("The factory interface for the Ribbon ModelFactory.")]
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IViewModelFactory)]
    public interface IViewModelFactory {
        /// <summary>.</summary>
        /// <param name="controlId"></param>
        [DispId(1), Description("")]
        IControlStrings GetStrings(string controlId);

        /// <summary>.</summary>
        /// <param name="controlId"></param>
        [DispId(2), Description("")]
        IControlStrings2 GetStrings2(string controlId);

        /// <summary>.</summary>
        [DispId(3), Description("")]
        object LoadImage(string imageId);
    }
}
