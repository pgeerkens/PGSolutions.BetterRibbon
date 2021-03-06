﻿////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;

using PGSolutions.RibbonUtilities.LinksAnalysis.Interfaces;

namespace PGSolutions.RibbonUtilities.LinksAnalysis {
    using Workbook = Microsoft.Office.Interop.Excel.Workbook;
    using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

    /// <summary>TODO</summary>
    [ComVisible(true)]
    [CLSCompliant(false)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.ILinksAnalyzer)]
    public interface ILinksAnalyzer {
        /// <summary>Attaches an {IIntegerSource} to the specified DropDown control.</summary>
        /// <param name="controlId">The ID of the control to be attached to the specified data source.</param>
        /// <param name="strings">The text strings to be displayed for this control.</param>
        /// <returns></returns>
        [Description( "Attaches an {IIntegerSource} to the specified DropDown control." )]
        ILinksLexer NewLinksLexer(ISourceCellRef cellRef, string formula);

        /// <summary>.</summary>
        /// <param name="excel"></param>
        /// <param name="nameList"></param>
        /// <remarks>
        /// Worksheets named any of "Links Errors", "Linked Files" and "Links Analysis" will be
        /// excluded from tha analysis.
        /// </remarks>
        ILinksAnalysis NewExternalLinksWB(Workbook wb);

        /// <summary>.</summary>
        /// <param name="excel"></param>
        /// <param name="nameList"></param>
        ILinksAnalysis NewExternalLinksWS(Worksheet ws);

        /// <summary>.</summary>
        /// <param name="excel"></param>
        /// <param name="nameList"></param>
        ILinksAnalysis Parse(ISourceCellRef cellRef, string formula);
    }
}
