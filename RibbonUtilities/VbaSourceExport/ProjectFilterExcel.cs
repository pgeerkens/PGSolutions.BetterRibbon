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
    public interface IWorkbookProcessor {
        void DoOnOpenWorkbook(string wkbkFullName, Action<VBE.VBProject, string> action);
    }

    [CLSCompliant(false)]
    public sealed class ProjectFilterExcel : ProjectFilter  {
        public ProjectFilterExcel(WorkbookProcessor processor): this(processor, null, null) { }

        public ProjectFilterExcel(WorkbookProcessor processor, string description, string extensions)
        : base(description, extensions) 
        => Processor = processor;

        WorkbookProcessor Processor { get; }

        /// <inheritdoc/>
        public override void ExtractProjects(FileDialogSelectedItems items, bool destIsSrc) {
            if (items == null ) throw new ArgumentNullException(nameof(items));
            foreach (string selectedItem in items) { ExtractProject(selectedItem, destIsSrc); }
        }

        /// <summary>Exports modules from specified EXCEL workbook to an eponymous subdirectory.</summary>
        private void ExtractProject(string path, bool destIsSrc)
        => Processor.DoOnOpenWorkbook(path,
                (p, s) => ExtractProjectModules(p, CreateDirectory(path, destIsSrc)));

        /// <summary>Exports modules from specified EXCEL workbook to an eponymous subdirectory.</summary>
        public void ExtractOpenProject(Workbook workbook, bool destIsSrc)
        => ExtractProjectModules(workbook?.VBProject, CreateDirectory(workbook?.FullName, destIsSrc));
    }

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

