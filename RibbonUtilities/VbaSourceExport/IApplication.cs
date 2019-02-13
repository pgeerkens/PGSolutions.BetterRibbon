////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

namespace PGSolutions.RibbonUtilities.VbaSourceExport {
    using Workbook = Microsoft.Office.Interop.Excel.Workbook;

    /// <summary>.</summary>
    [CLSCompliant(false)]
    public interface IApplication {
        /// <summary>.</summary>
        void DoOnOpenWorkbook(string wkbkFullName, Action<Workbook> action);

        /// <summary>.</summary>
        bool     DisplayAlerts   { get; set; }

        /// <summary>.</summary>
        dynamic  StatusBar       { get; set; }
    }
}
