////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

namespace PGSolutions.RibbonUtilities.VbaSourceExport {
    using ProjectFilters = IReadOnlyList<ProjectFilter>;

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
            if (filter == null) throw new ArgumentNullException(nameof(filter));

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

        public ProjectFilters FillFilters(FileDialog fileDialog) {
            if (fileDialog == null) throw new ArgumentNullException(nameof(fileDialog));

            var list = GetFilters(new WorkbookProcessor(Application));
            foreach (var item in list) {
                fileDialog.Filters.Add(item.Description, item.Extensions);
            }
            return list;
        }

        static ProjectFilters GetFilters(WorkbookProcessor processor) {
            var filters = new List<ProjectFilter> {
                new ProjectFilterExcel(processor,
                        "MS-Excel Projects", "*.xlsm;*.xlsb;*.xlam;*.xls;*.xla")
            };
            if (AccessWrapper.IsAccessSupported) {
                filters.Add(new ProjectFilterAccess(
                        "MS-Access Projects", "*.accdb;*.accda;*.mdb;*.mda"));
            }
            return filters.AsReadOnly();
        }

        private void OnStatusAvailable(object sender, EventArgs<string> e)
        => StatusAvailable?.Invoke(this, e);
    }
}

