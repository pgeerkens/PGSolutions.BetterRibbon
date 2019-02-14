////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017-8 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonUtilities.VbaSourceExport;

namespace PGSolutions.BetterRibbon {
    using static RibbonDispatcher.ComClasses.Extensions;

    internal sealed class VbaSourceExportModel : AbstractRibbonGroupModel2{
        private static string UseSrcFolderToggleID    = "UseSrcFolderToggle";
        private static string SelectedProjectButtonID = "SelectedProjectButton";
        private static string CurrentProjectButtonID  = "CurrentProjectButton";

        public VbaSourceExportModel(List<KeyValuePair<string,RibbonGroupViewModel>> viewModels)
        : base(viewModels) {
            DestIsSrc = GetModel<RibbonCheckBox>(UseSrcFolderToggleID, UseSrcFolderToggled, true, true, false.ToggleImage());
            ExportSelectedModel = GetModel<RibbonButton>(SelectedProjectButtonID, ExportSelected, true, true, "SaveAll");
            ExportCurrentModel  = GetModel<RibbonButton>(CurrentProjectButtonID, ExportCurrent, true, true, "FileSaveAs");

            DestIsSrc.IsPressed = false;

            Invalidate();
        }

        public RibbonToggleModel DestIsSrc           { get; }
        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        public RibbonButtonModel ExportSelectedModel { get; }
        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        public RibbonButtonModel ExportCurrentModel  { get; }

        private void UseSrcFolderToggled(object sender, bool isPressed) {
            DestIsSrc.IsPressed = isPressed;
            foreach (var kvp in ViewModels) {
                kvp.Value.GetControl<RibbonButton>($"{SelectedProjectButtonID}{kvp.Key}")
                            .IsEnabled = ! DestIsSrc.IsPressed;
                kvp.Value.GetControl<RibbonToggleButton>($"{UseSrcFolderToggleID}{kvp.Key}")
                            .SetImageMso(DestIsSrc.IsPressed.ToggleImage());
                kvp.Value.Invalidate();
            }
        }

        /// <summary>Extracts VBA modules from current EXCEL workbook to a sibling directory.</summary>
        /// <param name="sender">The object that initiated the event.</param>
        /// <remarks>
        /// Requires that access to the VBA project object model be trusted (Macro Security).
        /// </remarks>
        private void ExportCurrent(object sender) {
            if (!IsProjectModelTrusted()) { return; }
            var securitySaved = Application.AutomationSecurity;
            Application.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;

            try {
                Application.Cursor = XlMousePointer.xlWait;
                Application.StatusBar = "Exporting VBA Source ...";

                ProjectFilterExcel.ExtractOpenProject(Application.ActiveWorkbook, DestIsSrc.IsPressed);
            }
            catch (IOException ex) { ex.Message.MsgBoxShow(CallerName()); }
            finally {
                Application.AutomationSecurity = securitySaved;
                Application.StatusBar = "Ready";

                Application.Cursor = XlMousePointer.xlDefault;
            }
        }

        /// <summary>Extracts VBA modules from a selected EXCEL workbook to a sibling directory.</summary>
        /// <param name="sender">The object that initiated the event.</param>
        /// <remarks>
        /// Requires that access to the VBA project object model be trusted (Macro Security).
        /// </remarks>
        private void ExportSelected(object sender) {
            if (!IsProjectModelTrusted()) { return; }
            var securitySaved = Application.AutomationSecurity;
            Application.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;

            try {
                var fd = Application.FileDialog[MsoFileDialogType.msoFileDialogFilePicker];
                fd.Title = "Select VBA Project(s) to Export From";
                fd.ButtonName = "Export";
                fd.AllowMultiSelect = true;
                fd.Filters.Clear();
                fd.InitialFileName = Application.ActiveWorkbook?.Path ?? "C:\\";

                var list = new ProjectFilters(new WorkbookProcessor(Application));
                foreach (var item in list) {
                    fd.Filters.Add(item.Description, item.Extensions);
                }
                if (fd.Show() != 0) {
                    list[fd.FilterIndex-1].ExtractProjects(fd.SelectedItems, DestIsSrc.IsPressed);
                }
            }
            catch (IOException ex) { ex.Message.MsgBoxShow(CallerName()); }
            finally {
                Application.AutomationSecurity = securitySaved;
            }
        }

        private static bool IsProjectModelTrusted() {
            try { return Application.VBE != null; }
            catch (COMException) { PleaseEnableTrust(); }
            catch (InvalidOperationException) { PleaseEnableTrust(); }
            return false;
        }

        private static void PleaseEnableTrust()
        => "Please enable trust of the Project Object Model".MsgBoxShow("Project Model Not Trusted");

        private static Application Application => Globals.ThisAddIn.Application;
    }
}
