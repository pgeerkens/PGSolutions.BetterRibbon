////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;

using Microsoft.Office.Core;

namespace PGSolutions.RibbonUtilities {
    using Excel = Microsoft.Office.Interop.Excel;

    /// <summary>An implementation of <see cref="IWorkbookProcessor"/> that uses an existing <see cref="Excel.Application"/>.</summary>
    [CLSCompliant(false)]
    public class WorkbookProcessor : IWorkbookProcessor {
        /// <summary>.</summary>
        /// <param name="excelApp"></param>
        internal WorkbookProcessor(Excel.Application excelApp) => ExcelFG = excelApp;

        private Excel.Application ExcelFG  { get; }

        /// <inheritdoc/>
        public void DoOnWorkbook(string wkbkFullName, Action<Excel.Workbook> action) {
            Excel.Workbook wkbk = null;

            if( (wkbk = ExcelFG.Workbooks.TryItem(wkbkFullName))  !=  null) {
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
            var thisWkbk = ExcelFG.ActiveWorkbook;

            var saveSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;
            ExcelFG.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;
            ExcelFG.ScreenUpdating = false;
            ExcelFG.DisplayAlerts = false;

            Excel.Workbook wkbk = null;
            try {
                wkbk = ExcelFG.Workbooks.Open(wkbkFullName, UpdateLinks:false, ReadOnly:true,
                            AddToMru:false, Editable:false);
                
                action?.Invoke(wkbk);
            }
            finally {
                wkbk?.Close(false);

                ExcelFG.DisplayAlerts = true;
                ExcelFG.ScreenUpdating = true;
                ExcelFG.AutomationSecurity = saveSecurity;

                thisWkbk.Activate();
            }
        }

        public static IWorkbookProcessor New(Excel.Application application) => New(application, false);

        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope")]
        public static IWorkbookProcessor New(Excel.Application application, bool inBackground)
            => inBackground ? new WorkbookProcessor2(application)
                            : new WorkbookProcessor(application);

        #region Standard IDisposable baseclass implementation w/ Finalizer
        private bool _isDisposed = false;

        public void Dispose() { Dispose(true); GC.SuppressFinalize(this); }

        protected virtual void Dispose(bool disposing) {
            if (!_isDisposed) {

                // Dispose of managed resources (only!) here
                if (disposing) {
                }

                // Dispose of unmanaged resources here

                // Indicate that the instance has been disposed.
                _isDisposed = true;
            }
        }
        #endregion

        /// <summary>An implementation of <see cref="IWorkbookProcessor"/> that uses a private instance of Excel. <see cref="Excel.Application"/>.</summary>
        private sealed class WorkbookProcessor2:WorkbookProcessor {
            /// <summary>Creates and returns a new instance of <see cref="WorkbookProcessor2"/>.</summary>
            public WorkbookProcessor2(Excel.Application excelApp) : base(excelApp) => ExcelBG = NewExcelApp;

            private Excel.Application ExcelBG { get; }

            /// <inheritdoc/>
            protected override void DoOnClosedWorkbook(string wkbkFullName, Action<Excel.Workbook> action) {
                Excel.Workbook wkbk = null;
                try {
                    wkbk = ExcelBG.Workbooks.Open(wkbkFullName, UpdateLinks: false, ReadOnly: true,
                                        AddToMru: false, Editable: false);
                    ExcelBG.ActiveWindow.Visible = false;

                    DoOnOpenWorkbook(wkbk, action);
                }
                finally {
                    wkbk?.Close(false);
                }
            }

            /// <summary>Returns a new, invisible nad quiet, instance of <see cref="Excel.Application"/>.</summary>
            private static Excel.Application NewExcelApp
            => new Excel.Application {
                AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable,
                ScreenUpdating = false,
                DisplayAlerts = false
            };

            #region Standard IDisposable subclass implementation w/ Finalizer
            private new bool _isDisposed = false;

            protected override void Dispose(bool disposing) {
                if (!_isDisposed) {

                    // Dispose of managed resources (only!) here
                    if (disposing) {
                        ExcelBG?.Quit();
                    }

                    // Dispose of unmanaged resources here

                    // Indicate that the instance has been disposed.
                    _isDisposed = true;
                }
                base.Dispose(disposing);
            }
            #endregion
        }
    }
}

