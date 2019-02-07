////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;

using Microsoft.Office.Core;

namespace PGSolutions.RibbonUtilities.VbaSourceExport {
    using Application = Microsoft.Office.Interop.Excel.Application;
    
    [CLSCompliant(false)]
    public class ProjectFilterExcel : ProjectFilter  {
        public ProjectFilterExcel(IApplication application)
        : this(application, null, null) { }

        public ProjectFilterExcel(IApplication application, string description, string extensions)
        : base(application, description, extensions) { }

        /// <inheritdoc/>
        public override void ExtractProjects(FileDialogSelectedItems items, bool destIsSrc) {
            var app = new Lazy<Application>(() => new Application());
            try {
                foreach (string selectedItem in items) {
                    ExtractProject(app, selectedItem, destIsSrc);
                }
            } finally {
                if (app.IsValueCreated) { app.Value.Quit(); }
            }
        }

        /// <summary>Exports modules from specified EXCEL workbook to an eponymous subdirectory.</summary>
        private void ExtractProject(Lazy<Application> app, string filename, bool destIsSrc) {
            if (filename == Application.ActiveWorkbook.FullName) {
                ExtractOpenProject(Application.ActiveWorkbook, destIsSrc);
            } else {
                app.Value.Visible = false;
                app.Value.DisplayAlerts = false;
                app.Value.ScreenUpdating = false;
                app.Value.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;
                ExtractClosedProject(app.Value, filename, destIsSrc);
            }
        }

        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        private void ExtractClosedProject(Application app, string filename, bool destIsSrc) {
            var wkbk = app.Workbooks.Open(filename, UpdateLinks:false, ReadOnly:true, AddToMru:false, Editable:false);

            try {
                ExtractOpenProject(wkbk, destIsSrc);
            } finally {
                wkbk?.Close();
            }
        }

        ///// <inheritdoc/>
        //private void ExtractOpenProject(Excel.Workbook wkbk, bool destIsSrc)
        //=> ExtractProjectModules(wkbk.VBProject, CreateDirectory(wkbk.FullName, destIsSrc));
    }
}
