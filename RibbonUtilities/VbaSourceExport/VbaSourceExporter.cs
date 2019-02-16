////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

namespace PGSolutions.RibbonUtilities.VbaSourceExport {
    [CLSCompliant(false)]
    public class VbaSourceExporter {
        public VbaSourceExporter(Application application) => Application = application;

        public event EventHandler<EventArgs<string>> StatusAvailable;

        private Application Application { get; }

        public void ExtractOpenProject(_Workbook workbook, bool destIsSrc) {
            ProjectFilter.StatusAvailable += OnStatusAvailable;
            ProjectFilterExcel.ExtractOpenProject(workbook, destIsSrc);
            ProjectFilter.StatusAvailable -= OnStatusAvailable;
        }

        public void ExportSelected(ProjectFilter filter, FileDialogSelectedItems items, bool destIsSrc) {
            var securitySaved = Application.AutomationSecurity;
            try {
                ProjectFilter.StatusAvailable += OnStatusAvailable;
                Application.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;
                filter.ExtractProjects(items, destIsSrc);
            }
            finally {
                Application.AutomationSecurity = securitySaved;
                ProjectFilter.StatusAvailable -= OnStatusAvailable;
            }
        }

        public ProjectFilters FillFilters(FileDialog fd) {
            var list = new ProjectFilters(new WorkbookProcessor(Application));
            foreach (var item in list) {
                fd.Filters.Add(item.Description, item.Extensions);
            }
            return list;
        }

        private void OnStatusAvailable(object sender, EventArgs<string> e)
        => StatusAvailable?.Invoke(this, e);
    }
}

