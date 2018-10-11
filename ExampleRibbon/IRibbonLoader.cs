////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.AbstractCOM;
using System.ComponentModel;
using Microsoft.Office.Core;

namespace PGSolutions.SampleRibbon {
    /// <summary>The publicly available entry points to the library.</summary>
    [ComVisible(true)]
    [CLSCompliant(false)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IMain)]
    [Description("")]
    public interface IRibbonLoader {
        void InitializeRibbon(IRibbonUI ribbonUI);

        /// <summary>TODO</summary>
        [Description("")]
        void ReinitializeRibbon();

        /// <summary>TODO</summary>
        [Description("")]
        IRibbonViewModel RibbonViewModel { get; }
    }
}
