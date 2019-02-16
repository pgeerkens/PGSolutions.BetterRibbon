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
        public WorkbookProcessor(Application application) => Application = application;
        /// <inheritdoc/>
        public void DoOnOpenWorkbook(string wkbkFullName, Action<VBE.VBProject, string> action) {
            if (wkbkFullName == ActiveWorkbook.FullName) {
                action?.Invoke(ActiveWorkbook?.VBProject, Path.GetDirectoryName(wkbkFullName));
            } else {
                var thisWkbk = ActiveWorkbook;

                Application.DisplayAlerts = false;
                Application.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;

                Application.ScreenUpdating = false;
                var wkbk = Application.Workbooks.Open(wkbkFullName, UpdateLinks:false, ReadOnly:true,
                            AddToMru:false, Editable:false);
                Application.ActiveWindow.Visible = false;
                thisWkbk.Activate();

                try {
                    Application.ScreenUpdating = true;

                    action?.Invoke(wkbk?.VBProject, Path.GetDirectoryName(wkbkFullName));
                }
                finally {
                    wkbk?.Close(false);

                    Application.DisplayAlerts = true;
                }
            }
        }

        private Application Application { get; }

        private Workbook ActiveWorkbook => Application.ActiveWorkbook;
    }
}

