////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.IO;

using Microsoft.Office.Core;

namespace PGSolutions.RibbonUtilities {
    using Excel = Microsoft.Office.Interop.Excel;

    /// <summary>An implementation of <see cref="IWorkbookProcessor"/> that uses an existing <see cref="Excel.Application"/>.</summary>
    internal class WorkbookProcessor : IWorkbookProcessor {
        /// <summary>.</summary>
        /// <param name="excelApp"></param>
        public WorkbookProcessor(Excel.Application excelApp) => ExcelApp = excelApp;

        protected Excel.Application ExcelApp  { get; }

        private Excel.Workbook ActiveWorkbook => ExcelApp.ActiveWorkbook;

        /// <inheritdoc/>
        public void DoOnWorkbook(string wkbkFullName, Action<Excel.Workbook, string> action) {
            var path = Path.GetDirectoryName(wkbkFullName);
            var wkbk = ActiveWorkbook;

            if (wkbk.FullName == wkbkFullName) {
                action?.Invoke(ActiveWorkbook, path);
            } else if( (wkbk = ExcelApp.Workbooks.TryItem(wkbkFullName))  !=  null) {
                action?.Invoke(wkbk, path);
            } else {
                DoOnClosedWorkbook(wkbkFullName, action);
            }
        }

        /// <inheritdoc/>
        public virtual void DoOnClosedWorkbook(string wkbkFullName, Action<Excel.Workbook, string> action) {
            var path = Path.GetDirectoryName(wkbkFullName);
            var thisWkbk = ActiveWorkbook;

            var saveSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;
            ExcelApp.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;
            ExcelApp.ScreenUpdating = false;
            ExcelApp.DisplayAlerts = false;

            Excel.Workbook wkbk = null;
            try {
                ExcelApp.ActiveWindow.Visible = false;

                wkbk = ExcelApp.Workbooks.Open(wkbkFullName, UpdateLinks:false, ReadOnly:true,
                            AddToMru:false, Editable:false);
                action?.Invoke(wkbk, path);
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

