////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using System.Windows.Forms;

using Microsoft.Office.Core;
using Excel    = Microsoft.Office.Interop.Excel;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;

namespace PGSolutions.RibbonUtilities.VbaSourceExport {
    internal class ProjectFilterExcel : ProjectFilter  {
        internal ProjectFilterExcel(Excel.Application application)
        : this(application, "", "") { }

        public ProjectFilterExcel(Excel.Application application, string description, string extensions)
        : base(application, description, extensions) { }

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "System.Windows.Forms.MessageBox.Show(System.String,System.String,System.Windows.Forms.MessageBoxButtons,System.Windows.Forms.MessageBoxIcon,System.Windows.Forms.MessageBoxDefaultButton,System.Windows.Forms.MessageBoxOptions)")]
        public override void ExtractProjects(FileDialogSelectedItems items, bool destIsSrc) {
            if ( IsProjectModelTrusted) {
                foreach (string selectedItem in items) {
                    ExtractProject(selectedItem, destIsSrc);
                }
            } else {
                MessageBox.Show("Please enable trust of the Project Object Model", "Project Model Not Trusted",
                        MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
            }
        }

        /// <summary>Returns true exactly when the Project Object Model is trusted.</summary>
        private bool IsProjectModelTrusted => Application.VBE != null;

        /// <summary>Exports modules from specified EXCEL workbook to an eponymous subdirectory.</summary>
        private void ExtractProject(string filename, bool destIsSrc) {
            var appOpen   = Application;
            var appClosed = new Lazy<Excel.Application>(() => new Excel.Application());
            try {
                if (filename == appOpen.ActiveWorkbook.FullName) {
                    ExtractOpenProject(appOpen.ActiveWorkbook, destIsSrc);
                } else {
                    appClosed.Value.Visible = false;
                    appClosed.Value.DisplayAlerts = false;
                    appClosed.Value.ScreenUpdating = false;
                    appClosed.Value.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;
                    ExtractClosedProject(appClosed.Value, filename, destIsSrc);
                }
            } finally {
                if (appClosed.IsValueCreated) { appClosed.Value.Quit(); }
            }
        }

        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        private void ExtractClosedProject(Excel.Application app, string filename, bool destIsSrc) {
            var wkbk = app.Workbooks.Open(filename, UpdateLinks:false, ReadOnly:true, AddToMru:false, Editable:false);

            try {
                ExtractOpenProject(wkbk, destIsSrc);
            } finally {
                wkbk?.Close();
            }
        }

        /// <summary>Exports modules from specified EXCEL workbook to an eponymous subdirectory.</summary>
        public void ExtractOpenProject(Workbook wkbk, bool destIsSrc) =>
            ExtractProjectModules(wkbk.VBProject, CreateDirectory(wkbk.FullName, destIsSrc));
    }

}
