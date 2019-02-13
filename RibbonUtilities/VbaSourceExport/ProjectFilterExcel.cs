////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

namespace PGSolutions.RibbonUtilities.VbaSourceExport {   
    [CLSCompliant(false)]
    public sealed class ProjectFilterExcel : ProjectFilter  {
        public ProjectFilterExcel(IApplication application)
        : this(application, null, null) { }

        public ProjectFilterExcel(IApplication application, string description, string extensions)
        : base(application, description, extensions) { }

        /// <inheritdoc/>
        public override void ExtractProjects(FileDialogSelectedItems items, bool destIsSrc) {
            if (items == null ) throw new ArgumentNullException(nameof(items));
            //var app = new Lazy<Application>(() => new Application());
            try {
                foreach (string selectedItem in items) {
                    ExtractProject(selectedItem, destIsSrc);
                }
            } finally {
                //if (app.IsValueCreated) { app.Value.Quit(); }
            }
        }

        /// <summary>Exports modules from specified EXCEL workbook to an eponymous subdirectory.</summary>
        public void ExtractOpenProject(Workbook workbook, bool destIsSrc)
        => ExtractProjectModules(workbook?.VBProject, CreateDirectory(workbook?.FullName, destIsSrc));

        /// <summary>Exports modules from specified EXCEL workbook to an eponymous subdirectory.</summary>
        private void ExtractProject(string path, bool destIsSrc)
        => Application.DoOnOpenWorkbook(path, wb => ExtractOpenProject(wb,destIsSrc));
    }
}
