////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

using Microsoft.Office.Interop.Excel;

namespace PGSolutions.RibbonUtilities {
    public static class Extensions {
        /// <summary>Returns the speccified <see cref="Workbook"/> exactly when it is already open.</summary>
        /// <param name="excel">The running instnce of Excel.</param>
        /// <param name="path">The absolute full=path and -name for the desired workbook.</param>
        internal static Workbook TryItem(this Workbooks workbooks, string fullName) {
            foreach (Workbook wb in workbooks) if (wb.FullName == fullName) return wb;
            return null;
        }

        /// <summary>.</summary>
        /// <param name="excelApp"></param>
        /// <param name="path"></param>
        internal static void AnalyzeClosedWorkbook(this Application excelApp, string path,
                Action<Workbook> action) {
            Workbook wb = null;
            try {
                wb = excelApp.Workbooks.Open(path, UpdateLinks: false, ReadOnly: true, AddToMru: false);
                action(wb);
            }
            finally {
                wb?.Close(SaveChanges: false);
            }
        }
    }
}
