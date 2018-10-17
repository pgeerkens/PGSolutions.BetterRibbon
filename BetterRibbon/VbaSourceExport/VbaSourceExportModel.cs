////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017-8 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Collections.Generic;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

namespace PGSolutions.BetterRibbon.VbaSourceExport {
    internal sealed class VbaSourceExportModel {
        internal VbaSourceExportModel(IList<IVbaSourceExportGroupModel> viewModels ) {
            DestIsSrc  = false;
            ViewModels = viewModels;
            foreach (var viewModel in ViewModels) {
                viewModel.SelectedProjectsClicked += ExportSelectedProject;
                viewModel.CurrentProjectClicked   += ExportCurrentProject;
                viewModel.UseSrcFolderToggled     += UseSrcFolderToggled;
                viewModel.Attach(()=>DestIsSrc);
            }
        }

        private bool                              DestIsSrc  { get; set; }
        private IList<IVbaSourceExportGroupModel> ViewModels { get; set; }

        private void UseSrcFolderToggled(bool isPressed) {
            DestIsSrc = isPressed;
            foreach (var viewModel in ViewModels) {  viewModel.Invalidate(); }
        }

        /// <summary>Extracts VBA modules from current EXCEL workbook to a sibling directory.</summary>
        /// <param name="destIsSrc"> If true writes output to 'src'; else to a directory eponymous with the workbook.</param>
        /// <remarks>
        /// Requires that access to the VBA project object model be trusted (Macro Security).
        /// </remarks>
        private void ExportCurrentProject() => SilentAction(() => 
                ProjectFilterExcel.ExtractOpenProject(Application.ActiveWorkbook, DestIsSrc));

        /// <summary>Extracts VBA modules from a selected EXCEL workbook to a sibling directory.</summary>
        /// <param name="destIsSrc"> If true writes output to 'src'; else to a directory eponymous with the workbook.</param>
        /// <remarks>
        /// Requires that access to the VBA project object model be trusted (Macro Security).
        /// </remarks>
        private void ExportSelectedProject() {
            var securitySaved = Application.AutomationSecurity;
            Application.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;

            try {
                var fd = Application.FileDialog[MsoFileDialogType.msoFileDialogFilePicker];
                fd.AllowMultiSelect = !DestIsSrc;   // MultiSelect requires eponymous naming
                fd.ButtonName = "Export";
                fd.Title = "Select VBA Workbook(s) to Export From";
                fd.Filters.Clear();

                var list = new ProjectFilters();
                foreach (var item in list) {
                    fd.Filters.Add(item.Description, item.Extensions);
                }
                 if (fd.Show() != 0) {
                    SilentAction(
                        () => list[fd.FilterIndex].ExtractProjects(fd.SelectedItems, DestIsSrc)
                    );
                }

            } finally {
                Application.DisplayAlerts = true;
                Application.ScreenUpdating = true;
                Application.Cursor = XlMousePointer.xlDefault;
                Application.AutomationSecurity = securitySaved;
            }
        }

        private static void SilentAction(System.Action action) {
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
