////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017-8 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using Microsoft.Office.Interop.Excel;
using PGSolutions.RibbonDispatcher.ControlMixins;
using Microsoft.Office.Core;

namespace PGSolutions.ExcelRibbon.VbaSourceExport {
    internal static class VbaSourceExportModel {

        /// <summary>Extracts VBA modules from current EXCEL workbook to the sibling directory 'src'.</summary>
        /// <remarks>
        /// Requires that access to the VBA project object model be trusted (Macro Security).
        /// </remarks>
        public static void ExportCurrentProject() => ExportCurrentProject(true);

        /// <summary>Extracts VBA modules from current EXCEL workbook to a sibling directory.</summary>
        /// <param name="destIsSrc"> If true writes output to 'src'; else to a directory eponymous with the workbook.</param>
        /// <remarks>
        /// Requires that access to the VBA project object model be trusted (Macro Security).
        /// </remarks>
        public static void ExportCurrentProject(bool destIsSrc) {
            try {
                Globals.ThisAddIn.Application.Cursor = XlMousePointer.xlWait;
                Globals.ThisAddIn.Application.ScreenUpdating = false;
                ProjectFilterExcel.ExtractOpenProject(Globals.ThisAddIn.Application.ActiveWorkbook, destIsSrc);
            } finally {
                Globals.ThisAddIn.Application.StatusBar = false;
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                Globals.ThisAddIn.Application.Cursor = XlMousePointer.xlDefault;
            }

        }

        /// <summary>Extracts VBA modules from a selected EXCEL workbook to a sibling directory.</summary>
        /// <param name="destIsSrc"> If true writes output to 'src'; else to a directory eponymous with the workbook.</param>
        /// <remarks>
        /// Requires that access to the VBA project object model be trusted (Macro Security).
        /// </remarks>
        public static void ExportSelectedProject(bool destIsSrc) {
            var securitySaved = Globals.ThisAddIn.Application.AutomationSecurity;
            Globals.ThisAddIn.Application.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;

            try {
                var fd = Globals.ThisAddIn.Application.FileDialog[MsoFileDialogType.msoFileDialogFilePicker];
                fd.AllowMultiSelect = ! destIsSrc;   // MultiSelect requires eponymous naming
                fd.ButtonName = "Export";
                fd.Title = "Select VBA Workbook(s) to Export From";
                fd.Filters.Clear();

                var list = new ProjectFilters();
                foreach (var item in list) {
                    fd.Filters.Add(item.Description, item.Extensions);
                }
                 if (fd.Show() != 0) {
                    Globals.ThisAddIn.Application.Cursor = XlMousePointer.xlWait;
                    Globals.ThisAddIn.Application.ScreenUpdating = false;
                    Globals.ThisAddIn.Application.DisplayAlerts = false;
                    list[fd.FilterIndex].ExtractProjects(fd.SelectedItems, destIsSrc);
                }

            } finally {
                Globals.ThisAddIn.Application.DisplayAlerts = true;
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                Globals.ThisAddIn.Application.Cursor = XlMousePointer.xlDefault;
                Globals.ThisAddIn.Application.AutomationSecurity = securitySaved;
            }
        }
    }
}
