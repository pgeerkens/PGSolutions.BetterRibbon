////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    /// <summary>The factory interface for the Ribbon ModelFactory.</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IRibbonFactory)]
    public interface IRibbonFactory {
        /// <summary>TODO</summary>
        IResourceManager ResourceManager { get; }

        /// <summary>.</summary>
        /// <param name="controlId"></param>
        [Description("")]
        IControlStrings GetStrings(string controlId);

        /// <summary>Returns a new {ResourceLoader} object.</summary>
        IResourceLoader NewResourceLoader();

        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        IControlStrings NewControlStrings(string label,
                string screenTip = "", string superTip = "",
                string keyTip = "", string alternateLabel = "", string description = "");

        object LoadImage(string imageId);
    }
}
