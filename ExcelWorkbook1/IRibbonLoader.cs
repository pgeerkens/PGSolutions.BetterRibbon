////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.AbstractCOM;
using System.ComponentModel;

namespace PGSolutions.SampleRibbon {
    /// <summary>The publicly available entry points to the library.</summary>
    [ComVisible(true)]
    [CLSCompliant(false)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IMain)]
    [Description("")]
    public interface IRibbonLoader {
        /// <summary>TODO</summary>
        [Description("")]
        void ReinitializeRibbon();

        /// <summary>TODO</summary>
        [Description("")]
        IRibbonViewModel RibbonViewModel { get; }
    }
}
