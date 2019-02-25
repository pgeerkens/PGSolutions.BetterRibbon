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
    [CLSCompliant(false)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IViewModelFactory)]
    public interface IViewModelFactory {
        /// <summary>.</summary>
        /// <param name="controlId"></param>
        [DispId(1), Description("")]
        IControlStrings GetStrings(string controlId);

        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [DispId(2), Description("")]
        IControlStrings NewControlStrings(string label,
                                          string screenTip      = null,
                                          string superTip       = null,
                                          string keyTip         = null,
                                          string alternateLabel = null,
                                          string description    = null);

        /// <summary>.</summary>
        [DispId(3), Description("")]
        object LoadImage(string imageId);
    }
}
