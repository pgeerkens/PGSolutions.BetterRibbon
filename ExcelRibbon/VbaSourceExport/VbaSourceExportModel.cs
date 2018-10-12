////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017-8 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Collections.Generic;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

namespace PGSolutions.ExcelRibbon.VbaSourceExport {
    internal sealed class VbaSourceExportModel {
        internal VbaSourceExportModel(IList<IVbaSourceExportGroupModel> models ) {
            DestIsSrc = true;
            Models    = models;
            foreach (var model in Models) {
                model.SelectedProjectsClicked += ExportSelectedProject;
                model.CurrentProjectClicked   += ExportCurrentProject;
                model.UseSrcFolderToggled     += UseSrcFolderToggled;
                model.Attach(()=>DestIsSrc);
            }
        }

        private bool                              DestIsSrc { get; set; }
        private IList<IVbaSourceExportGroupModel> Models    { get; set; }

        private void UseSrcFolderToggled(bool isPressed) => DestIsSrc = isPressed;

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

        private void SilentAction(System.Action action) {
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

        private Application Application => Globals.ThisAddIn.Application;
    }
}
