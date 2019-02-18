////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using Microsoft.Office.Core;

namespace PGSolutions.RibbonUtilities {
    using Excel = Microsoft.Office.Interop.Excel;
    using Marshal = System.Runtime.InteropServices.Marshal;

    /// <summary>An implementation of <see cref="IWorkbookProcessor"/> that uses a private instance of Excel. <see cref="Excel.Application"/>.</summary>
    [CLSCompliant(false)]
    public sealed class WorkbookProcessor2 : WorkbookProcessor, IDisposable {
        /// <summary>Creates and returns a new instance of <see cref="WorkbookProcessor2"/>.</summary>
        public WorkbookProcessor2(Excel.Application excelApp) : base(excelApp) => ExcelApp = NewExcelApp();

        private new Excel.Application ExcelApp { get; }

        /// <inheritdoc/>
        protected override void DoOnClosedWorkbook(string wkbkFullName, Action<Excel.Workbook> action){
            Excel.Workbook wkbk = null;
            try {
                wkbk = ExcelApp.Workbooks.Open(wkbkFullName, UpdateLinks:false, ReadOnly:true,
                                    AddToMru:false, Editable:false);
                ExcelApp.ActiveWindow.Visible = false;

                DoOnOpenWorkbook(wkbk, action);
            }
            finally {
                wkbk?.Close(false);
                Marshal.ReleaseComObject(wkbk);
            }
        }

        /// <summary>Returns a new, invisible nad quiet, instance of <see cref="Excel.Application"/>.</summary>
        private static Excel.Application NewExcelApp()
        => new Excel.Application {
                AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable,
                ScreenUpdating = false,
                DisplayAlerts = false
        };

        private static bool Use2 => true;

        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope")]
        public static IWorkbookProcessor New(Excel.Application application)
            => Use2 ? new WorkbookProcessor2(application)
                    : new WorkbookProcessor(application);

        #region Standard IDisposable baseclass implementation w/ Finalizer
        private bool _isDisposed = false;

        public void Dispose() { Dispose(true); GC.SuppressFinalize(this); }

        private void Dispose(bool disposing) {
            if (!_isDisposed) {

                // Dispose of managed resources (only!) here
                if (disposing) {
                    if (ExcelApp != null) {
                        ExcelApp.Quit();
                        Marshal.ReleaseComObject(ExcelApp.Workbooks);
                        Marshal.ReleaseComObject(ExcelApp);
                    }
                }

                // Dispose of unmanaged resources here

                // Indicate that the instance has been disposed.
                _isDisposed = true;
            }
        }
        #endregion
    }
}

