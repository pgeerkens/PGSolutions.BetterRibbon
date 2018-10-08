////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017-8 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using Microsoft.Office.Interop.Excel;
using PGSolutions.RibbonDispatcher2013.ControlMixins;
using Microsoft.Office.Core;

namespace PGSolutions.ExcelRibbon2013 {
    internal static class ExportVba {

        /// <summary>Exports all VBA modules in the current workbook to a sibling directory named 'src'.</summary>
        /// <remarks>
        /// The module files are saved in a subdirectory 'src'
        ///
        /// Requires that access to the VBA project object model be trusted (Macro Security).
        /// </remarks>
        public static ClickedEventHandler ExportVbaModulesCurrent() => () => ExportModulesCurrentProject();

        /// <summary>Exports all VBA modules in a selected workbook to eponymous files.</summary>
        /// <remarks>
        /// The module files are saved in a subdirectory 'src'.
        ///
        /// Requires that access to the VBA project object model be trusted (Macro Security).
        /// </remarks>
        public static ClickedEventHandler ExportVbaModules() => () => ExportModules(false);

        /// <summary>Extracts VBA modules from current EXCEL workbook to the sibling directory 'src'.</summary>
        public static void ExportModulesCurrentProject() => ExportModulesCurrentProject(true);

        /// <summary>Extracts VBA modules from current EXCEL workbook to a sibling directory.</summary>
        /// <param name="destIsSrc"> If true writes output to 'src'; else to a directory eponymous with the workbook.</param>
        public static void ExportModulesCurrentProject(bool destIsSrc) {
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
        public static void ExportModules(bool destIsSrc) {
            try {
                var fd = Globals.ThisAddIn.Application.FileDialog[MsoFileDialogType.msoFileDialogFilePicker];
                fd.AllowMultiSelect = false;
                fd.ButtonName = "Export";
                fd.Title = "Select VBA Project(s) to Export From";
                fd.Filters.Clear();

                var list = new ProjectFilters();
                foreach (var item in list) {
                    fd.Filters.Add(item.Description, item.Extensions);
                }
                 if (fd.Show() != 0) {
                    Globals.ThisAddIn.Application.Cursor = XlMousePointer.xlWait;
                    Globals.ThisAddIn.Application.ScreenUpdating = false;
                    list[fd.FilterIndex].ExtractProjects(fd.SelectedItems, destIsSrc);
                }

            } finally {
                Globals.ThisAddIn.Application.StatusBar = false;
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                Globals.ThisAddIn.Application.Cursor = XlMousePointer.xlDefault;
            }
        }
        
    }
}
