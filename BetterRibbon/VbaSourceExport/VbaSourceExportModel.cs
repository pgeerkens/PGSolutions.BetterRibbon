////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017-8 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Collections.Generic;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.BetterRibbon.VbaSourceExport {
    internal sealed class VbaSourceExportModel : IBooleanSource {
        internal VbaSourceExportModel(IList<IVbaSourceExportViewModel> viewModels ) {
            DestIsSrc  = false;
            ViewModels = viewModels;
            foreach (var viewModel in ViewModels) {
                viewModel.SelectedProjectsClicked += ExportSelectedProject;
                viewModel.CurrentProjectClicked   += ExportCurrentProject;
                viewModel.UseSrcFolderToggled     += UseSrcFolderToggled;
                viewModel.Attach(this);
            }
        }

        bool IBooleanSource.Getter() => DestIsSrc;

        /// <summary>Fakse => file destination is eponymous directory; else directory named "SRC".</summary>
        private bool                             DestIsSrc  { get; set; }
        private IList<IVbaSourceExportViewModel> ViewModels { get; }

        private void UseSrcFolderToggled(object sender, bool isPressed) {
            DestIsSrc = isPressed;
            foreach (var viewModel in ViewModels) {
                viewModel.SelectedProjectButton.IsEnabled = ! DestIsSrc;
                viewModel.Invalidate();
            }
        }

        /// <summary>Extracts VBA modules from current EXCEL workbook to a sibling directory.</summary>
        /// <param name="sender">The object that initiated the event.</param>
        /// <remarks>
        /// Requires that access to the VBA project object model be trusted (Macro Security).
        /// </remarks>
        private void ExportCurrentProject(object sender) => PerformSilently(() => 
                ProjectFilterExcel.ExtractOpenProject(Application.ActiveWorkbook, DestIsSrc));

        /// <summary>Extracts VBA modules from a selected EXCEL workbook to a sibling directory.</summary>
        /// <param name="sender">The object that initiated the event.</param>
        /// <remarks>
        /// Requires that access to the VBA project object model be trusted (Macro Security).
        /// </remarks>
        private void ExportSelectedProject(object sender) {
            var securitySaved = Application.AutomationSecurity;
            Application.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;

            try {
                var fd = Application.FileDialog[MsoFileDialogType.msoFileDialogFilePicker];
                fd.AllowMultiSelect = ! DestIsSrc;   // MultiSelect requires eponymous naming
                fd.ButtonName = "Export";
                fd.Title = "Select VBA Project(s) to Export From";
                fd.Filters.Clear();
                fd.InitialFileName = Application.ActiveWorkbook?.Path ?? "C:\\";

                var list = new ProjectFilters();
                foreach (var item in list) {
                    fd.Filters.Add(item.Description, item.Extensions);
                }
                if (fd.Show() != 0) {
                    PerformSilently(
                        () => list[fd.FilterIndex-1].ExtractProjects(fd.SelectedItems, DestIsSrc)
                    );
                }
            } finally {
                Application.DisplayAlerts = true;
                Application.ScreenUpdating = true;
                Application.Cursor = XlMousePointer.xlDefault;
                Application.AutomationSecurity = securitySaved;
            }
        }

        private static void PerformSilently(System.Action action) {
            try {
                Application.Cursor = XlMousePointer.xlWait;
                Application.ScreenUpdating = false;
                Application.DisplayAlerts = false;
                action();
            } finally {
                Application.StatusBar = false;

                Application.DisplayAlerts = true;
                Application.ScreenUpdating = true;
                Application.Cursor = XlMousePointer.xlDefault;
            }
        }

        private static Application Application => Globals.ThisAddIn.Application;
    }
}
