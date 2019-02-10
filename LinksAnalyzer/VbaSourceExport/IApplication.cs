////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

using Microsoft.Office.Core;

namespace PGSolutions.RibbonUtilities.VbaSourceExport {
    using Workbook = Microsoft.Office.Interop.Excel.Workbook;

    /// <summary>.</summary>
    [CLSCompliant(false)]
    public interface IApplication {
        /// <summary>.</summary>
        Workbook ActiveWorkbook  { get; }

        /// <summary>.</summary>
        bool     DisplayAlerts   { get; set; }

        /// <summary>.</summary>
        dynamic  StatusBar       { get; set; }

        /// <summary>.</summary>
        MsoAutomationSecurity AutomationSecurity { get; set; }
    }
}
