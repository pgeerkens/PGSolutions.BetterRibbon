////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

namespace PGSolutions.RibbonUtilities {
    using Excel = Microsoft.Office.Interop.Excel;

    /// <summary>.</summary>
    internal interface IWorkbookProcessor {
        /// <summary>Performs the specified <paramref name="action"/> on <paramref name="wkbkFullName"/>,
        /// silently &amp; safely opening &amp; closing it as necessary.</summary>
        /// <param name="wkbkFullName">Full absolute path and name for the workbok to be acted upon.</param>
        /// <param name="action">The <see cref="Action"/> to be performed on the workbook.</param>
        void DoOnWorkbook(string wkbkFullName, Action<Excel.Workbook, string> action);

        /// <summary>Performs the specified <paramref name="action"/> on <paramref name="wkbkFullName"/>,
        /// silently &amp; safely opening &amp; closing it.</summary>
        /// <param name="wkbkFullName">Full absolute path and name for the workbok to be acted upon.</param>
        /// <param name="action">The <see cref="Action"/> to be performed on the workbook.</param>
        void DoOnClosedWorkbook(string wkbkFullName, Action<Excel.Workbook, string> action);
    }
}

