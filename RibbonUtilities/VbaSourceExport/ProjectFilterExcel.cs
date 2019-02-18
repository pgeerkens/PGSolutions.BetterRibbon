////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.IO;
using Microsoft.Office.Core;

using Microsoft.Office.Interop.Excel;

namespace PGSolutions.RibbonUtilities.VbaSourceExport {
    internal sealed class ProjectFilterExcel : ProjectFilter  {

        public ProjectFilterExcel(IWorkbookProcessor processor, string description, string extensions)
        : base(description, extensions) 
        => Processor = processor;

        IWorkbookProcessor Processor { get; }

        /// <inheritdoc/>
        public override void ExtractProjects(FileDialogSelectedItems items, bool destIsSrc) {
            if (items == null ) throw new ArgumentNullException(nameof(items));
            foreach (string selectedItem in items) {
                OnStatusAvailable(this, $"Exporting VBA Source from {selectedItem}; Please be patient ...");
                ExtractProject(selectedItem, destIsSrc);
            }
        }

        /// <summary>Exports modules from specified EXCEL workbook to an eponymous subdirectory.</summary>
        private void ExtractProject(string wkbkFullName, bool destIsSrc) {
            var path = Path.GetDirectoryName(wkbkFullName);
            path = wkbkFullName;
            Processor.DoOnWorkbook(wkbkFullName,
                    wb => ExtractProjectModules(wb?.VBProject, CreateDirectory(path,destIsSrc)));
        }

        /// <summary>Exports modules from specified EXCEL workbook to an eponymous subdirectory.</summary>
        internal static void ExtractOpenProject(Workbook workbook, bool destIsSrc) {
            OnStatusAvailable(workbook, $"Exporting VBA Source from {workbook.FullName}; Please be patient ...");
            ExtractProjectModules(workbook?.VBProject, CreateDirectory(workbook?.FullName, destIsSrc));
        }
    }
}

