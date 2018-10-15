﻿////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

using System.Runtime.InteropServices;

using System.ComponentModel;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary>The publicly available entry points to the library.</summary>
    [ComVisible(true)]
    [CLSCompliant(false)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IRibbonLoader)]
    [Description("The publicly available entry points to the library.")]
    public interface IRibbonLoader {
        /// <summary>Reinitializes the custom ribbon for this workbook from the cachedd {IRibbonUI}.</summary>
        /// <remarks>
        /// This is useful during code development after recompiles.
        /// </remarks>
        [Description("Reinitializes the custom ribbon for this workbook from the cachedd {IRibbonUI}.")]
        void ReinitializeRibbon();

        /// <summary>Returns the {IRibbonViewModel} for this workbook.</summary>
        [Description("Returns the {IRibbonViewModel} for this workbook.")]
        IRibbonViewModel RibbonViewModel { get; }
    }
}