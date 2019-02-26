////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary>The default COM interface exposed by {ResourceLoader} objects.</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IResourceLoader)]
    public interface IResourceLoader {
        /// <summary>Adds a new ControlString to the collection, and returns it.</summary>
        [DispId(1), Description("Returns a new ControlString from local resources.")]
        IControlStrings GetControlStrings(string ControlId);

        /// <summary>Adds a new ControlString to the collection, and returns it.</summary>
        [DispId(2), Description("Returns a new ControlString from local resources.")]
        IControlStrings2 GetControlStrings2(string ControlId);
    }
}
