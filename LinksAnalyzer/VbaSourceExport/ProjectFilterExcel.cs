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
    [CLSCompliant(false)]
    public class ProjectFilterExcel : ProjectFilter  {
        public ProjectFilterExcel(IApplication application)
        : this(application, null, null) { }

        public ProjectFilterExcel(IApplication application, string description, string extensions)
        : base(application, description, extensions) { }

        /// <inheritdoc/>
        public override void ExtractProjects(FileDialogSelectedItems items, bool destIsSrc) {
            foreach (string selectedItem in items) {
                ExtractProject(selectedItem, destIsSrc);
            }
        }

        /// <summary>Exports modules from specified EXCEL workbook to an eponymous subdirectory.</summary>
        private void ExtractProject(string filename, bool destIsSrc) {
            var appClosed = new Lazy<Excel.Application>(() => new Excel.Application());
            try {
                if (filename == Application.ActiveWorkbook.FullName) {
                    ExtractOpenProject(Application.ActiveWorkbook, destIsSrc);
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
    }

    /// <summary>.</summary>
    [CLSCompliant(false)]
    public interface IApplication {
        Workbook ActiveWorkbook         { get; }

        bool     DisplayAlerts          { get; set; }

        dynamic  StatusBar              { get; set; }

        MsoAutomationSecurity AutomationSecurity { get; set; }
    }

    public static partial class Extensions {
        /// <summary>.</summary>
        /// <param name="this"></param>
        internal static AccessWrapper NewAccessWrapper(this IApplication @this) => AccessWrapper.New(@this);
    }
}
