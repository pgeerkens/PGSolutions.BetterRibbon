////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

namespace PGSolutions.RibbonUtilities.VbaSourceExport {
    using VBE = Microsoft.Vbe.Interop;

    internal interface IWorkbookProcessor {
        void DoOnOpenWorkbook(string wkbkFullName, Action<VBE.VBProject, string> action);
    }

    internal sealed class ProjectFilterExcel : ProjectFilter  {

        public ProjectFilterExcel(WorkbookProcessor processor, string description, string extensions)
        : base(description, extensions) 
        => Processor = processor;

        WorkbookProcessor Processor { get; }

        /// <inheritdoc/>
        public override void ExtractProjects(FileDialogSelectedItems items, bool destIsSrc) {
            if (items == null ) throw new ArgumentNullException(nameof(items));
            foreach (string selectedItem in items) {
                OnStatusAvailable(this, $"Exporting VBA Source from {selectedItem}; Please be patient ...");
                ExtractProject(selectedItem, destIsSrc);
            }
        }

        /// <summary>Exports modules from specified EXCEL workbook to an eponymous subdirectory.</summary>
        private void ExtractProject(string path, bool destIsSrc)
        => Processor.DoOnOpenWorkbook(path,
                (p, s) => ExtractProjectModules(p, CreateDirectory(path, destIsSrc)));

        /// <summary>Exports modules from specified EXCEL workbook to an eponymous subdirectory.</summary>
        internal static void ExtractOpenProject(_Workbook workbook, bool destIsSrc) {
            OnStatusAvailable(workbook, $"Exporting VBA Source from {workbook.FullName}; Please be patient ...");
            ExtractProjectModules(workbook?.VBProject, CreateDirectory(workbook?.FullName, destIsSrc));
        }
    }
}

