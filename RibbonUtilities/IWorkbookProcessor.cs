////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

namespace PGSolutions.RibbonUtilities {
    using Excel = Microsoft.Office.Interop.Excel;

    /// <summary>.</summary>
    [CLSCompliant(false)]
    public interface IWorkbookProcessor : IDisposable {
        /// <summary>Performs the specified <paramref name="action"/> on <paramref name="wkbkFullName".</summary>
        /// <param name="wkbkFullName">Full absolute path and name for the workbok to be acted upon.</param>
        /// <param name="action">The <see cref="Action"/> to be performed on the workbook.</param>
        void DoOnWorkbook(string wkbkFullName, Action<Excel.Workbook> action);
    }
}

