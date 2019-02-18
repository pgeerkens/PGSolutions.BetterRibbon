////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.IO;

using Microsoft.Office.Core;

namespace PGSolutions.RibbonUtilities {
    using Excel = Microsoft.Office.Interop.Excel;

    /// <summary>An implementation of <see cref="IWorkbookProcessor"/> that uses an existing <see cref="Excel.Application"/>.</summary>
    [CLSCompliant(false)]
    public class WorkbookProcessor : IWorkbookProcessor {
        /// <summary>.</summary>
        /// <param name="excelApp"></param>
        internal WorkbookProcessor(Excel.Application excelApp) => ExcelApp = excelApp;

        protected Excel.Application ExcelApp  { get; }

        private Excel.Workbook ActiveWorkbook => ExcelApp.ActiveWorkbook;

        /// <inheritdoc/>
        public void DoOnWorkbook(string wkbkFullName, Action<Excel.Workbook> action) {
            Excel.Workbook wkbk = null;

            if( (wkbk = ExcelApp.Workbooks.TryItem(wkbkFullName))  !=  null) {
                DoOnOpenWorkbook(wkbk, action);
            } else {
                DoOnClosedWorkbook(wkbkFullName, action);
            }
        }

        /// <inheritdoc/>
        protected static void DoOnOpenWorkbook(Excel.Workbook wkbk, Action<Excel.Workbook> action)
        => action?.Invoke(wkbk);

        /// <inheritdoc/>
        protected virtual void DoOnClosedWorkbook(string wkbkFullName, Action<Excel.Workbook> action) {
            var thisWkbk = ActiveWorkbook;

            var saveSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;
            ExcelApp.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;
            ExcelApp.ScreenUpdating = false;
            ExcelApp.DisplayAlerts = false;

            Excel.Workbook wkbk = null;
            try {
            //    ExcelApp.ActiveWindow.Visible = false;

                wkbk = ExcelApp.Workbooks.Open(wkbkFullName, UpdateLinks:false, ReadOnly:true,
                            AddToMru:false, Editable:false);
                
                action?.Invoke(wkbk);
            }
            finally {
                wkbk?.Close(false);

                ExcelApp.DisplayAlerts = true;
                ExcelApp.ScreenUpdating = true;
                ExcelApp.AutomationSecurity = saveSecurity;

                thisWkbk.Activate();
            }
        }
    }
}

