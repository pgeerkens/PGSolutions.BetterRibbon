////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.IO;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

namespace PGSolutions.RibbonUtilities.VbaSourceExport {
    using VBE = Microsoft.Vbe.Interop;

    [CLSCompliant(false)]
    public sealed class WorkbookProcessor : IWorkbookProcessor {
        public WorkbookProcessor(Application excelApp) => ExcelApp = excelApp;
        /// <inheritdoc/>
        public void DoOnOpenWorkbook(string wkbkFullName, Action<VBE.VBProject, string> action) {
            Workbook wb = ActiveWorkbook;
            if (wb.FullName == wkbkFullName) {
                action?.Invoke(ActiveWorkbook?.VBProject, Path.GetDirectoryName(wkbkFullName));
            } else if( (wb = ExcelApp.Workbooks.TryItem(wkbkFullName))  !=  null) {
                action?.Invoke(wb.VBProject, Path.GetDirectoryName(wkbkFullName));
            } else {
                var thisWkbk = ActiveWorkbook;

                ExcelApp.DisplayAlerts = false;
                ExcelApp.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;

                ExcelApp.ScreenUpdating = false;
                var wkbk = ExcelApp.Workbooks.Open(wkbkFullName, UpdateLinks:false, ReadOnly:true,
                            AddToMru:false, Editable:false);
                ExcelApp.ActiveWindow.Visible = false;
                thisWkbk.Activate();

                try {
                    ExcelApp.ScreenUpdating = true;

                    action?.Invoke(wkbk?.VBProject, Path.GetDirectoryName(wkbkFullName));
                }
                finally {
                    wkbk?.Close(false);

                    ExcelApp.DisplayAlerts = true;
                }
            }
        }

        private Application ExcelApp { get; }

        private Workbook ActiveWorkbook => ExcelApp.ActiveWorkbook;
    }
}

