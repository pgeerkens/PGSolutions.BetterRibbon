////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017-8 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.ExcelRibbon.VbaSourceExport {
    internal sealed class VbaSourceExportModel {
        public VbaSourceExportModel(IRibbonFactory factory) {
            DestIsSrc = true;

            VbaExportGroupMS = new VbaSourceExportViewModel(factory, "MS", ()=>DestIsSrc);
            VbaExportGroupMS.SelectedProjectsClicked += ExportSelectedProject;
            VbaExportGroupMS.CurrentProjectClicked   += ExportCurrentProject;
            VbaExportGroupMS.UseSrcFolderToggled     += UseSrcFolderToggled;
            VbaExportGroupMS.SelectedProjectButton.Attach(null);
            VbaExportGroupMS.CurrentProjectButton.Attach(null);

            VbaExportGroupPG = new VbaSourceExportViewModel(factory, "PG", () => DestIsSrc);
            VbaExportGroupMS.SelectedProjectsClicked += ExportSelectedProject;
            VbaExportGroupMS.CurrentProjectClicked   += ExportCurrentProject;
            VbaExportGroupPG.UseSrcFolderToggled     += UseSrcFolderToggled;
            VbaExportGroupPG.SelectedProjectButton.Attach(null);
            VbaExportGroupPG.CurrentProjectButton.Attach(null);
        }

        public bool                     DestIsSrc        { get; private set; }
        public VbaSourceExportViewModel VbaExportGroupMS { get; private set; }
        public VbaSourceExportViewModel VbaExportGroupPG { get; private set; }

        public void UseSrcFolderToggled(bool isPressed) => DestIsSrc = isPressed;

        /// <summary>Extracts VBA modules from current EXCEL workbook to a sibling directory.</summary>
        /// <param name="destIsSrc"> If true writes output to 'src'; else to a directory eponymous with the workbook.</param>
        /// <remarks>
        /// Requires that access to the VBA project object model be trusted (Macro Security).
        /// </remarks>
        public void ExportCurrentProject() {
            try {
                Globals.ThisAddIn.Application.Cursor = XlMousePointer.xlWait;
                Globals.ThisAddIn.Application.ScreenUpdating = false;
                ProjectFilterExcel.ExtractOpenProject(Globals.ThisAddIn.Application.ActiveWorkbook, DestIsSrc);
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
        public void ExportSelectedProject() {
            var securitySaved = Globals.ThisAddIn.Application.AutomationSecurity;
            Globals.ThisAddIn.Application.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;

            try {
                var fd = Globals.ThisAddIn.Application.FileDialog[MsoFileDialogType.msoFileDialogFilePicker];
                fd.AllowMultiSelect = !DestIsSrc;   // MultiSelect requires eponymous naming
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
                    list[fd.FilterIndex].ExtractProjects(fd.SelectedItems, DestIsSrc);
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
