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
        /// <summary>Returns the specified {ControlStrings} object.</summary>
        [Description("Returns the specified ControlStrings object.")]
        IControlStrings this[string ControlId] { get; }

        /// <summary>Adds a new ControlString to the collection, and returns it.</summary>
        [Description("Adds a new ControlString to the collection, and returns it.")]
        string AddControlStrings(string ControlId,
            [Optional]string Label,
            [Optional]string ScreenTip,
            [Optional]string SuperTip,
            [Optional]string KeyTip);

        /// <summary>Adds a new ControlString to the collection, and returns it.</summary>
        [Description("Adds a new ControlString to the collection, and returns it.")]
        string AddControlStrings2(string ControlId,
            [Optional]string Label,
            [Optional]string ScreenTip,
            [Optional]string SuperTip,
            [Optional]string Description,
            [Optional]string KeyTip);
    }
}
