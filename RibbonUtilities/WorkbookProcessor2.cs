////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

using Microsoft.Office.Core;

namespace PGSolutions.RibbonUtilities {
    using Excel = Microsoft.Office.Interop.Excel;

    /// <summary>An implementation of <see cref="IWorkbookProcessor"/> that uses a private instance of Excel. <see cref="Excel.Application"/>.</summary>
    internal class WorkbookProcessor2 : WorkbookProcessor, IDisposable {
        /// <summary>Creates and returns a new instance of <see cref="WorkbookProcessor2"/>.</summary>
        /// <param name="excelApp"></param>
        public WorkbookProcessor2() : base(NewExcelApp()) { }

        /// <inheritdoc/>
        public override void DoOnClosedWorkbook(string wkbkFullName, Action<Excel.Workbook, string> action){
            Excel.Workbook wkbk = null;
            try {
                wkbk = ExcelApp.Workbooks.Open(wkbkFullName, UpdateLinks:false, ReadOnly:true,
                            AddToMru:false, Editable:false);
            }
            finally {
                wkbk?.Close(false);
            }
        }

        /// <summary>Returns a new, invisible nad quiet, instance of <see cref="Excel.Application"/>.</summary>
        private static Excel.Application NewExcelApp() {
            var excelApp = new Excel.Application();
            excelApp.ActiveWindow.Visible = false;
            excelApp.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;
            excelApp.ScreenUpdating = false;
            excelApp.DisplayAlerts = false;

            return excelApp;
        }

        #region Standard IDisposable baseclass implementation w/ Finalizer
        private bool _isDisposed = false;

        public void Dispose() { Dispose(true); GC.SuppressFinalize(this); }

        protected virtual void Dispose(bool disposing) {
            if (!_isDisposed) {

                // Dispose of managed resources (only!) here
                if (disposing) {
                    ExcelApp?.Quit();
                }

                // Dispose of unmanaged resources here

                // Indicate that the instance has been disposed.
                _isDisposed = true;
            }
        }
        #endregion
    }
}

