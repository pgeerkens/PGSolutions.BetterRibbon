////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017-8 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonUtilities.VbaSourceExport;

namespace PGSolutions.BetterRibbon {
    using static RibbonDispatcher.ComClasses.Extensions;
    using Models = List<VbaSourceExportGroupModel>;

    internal sealed class VbaSourceExportModel {

        public VbaSourceExportModel(Models models) {
            Models    = models;
            DestIsSrc = false;

            Models.ForEach(model => {
                model.UseSrcFolderToggled   += UseSrcFolderToggled;
                model.ExportSelectedClicked += ExportSelected;
                model.ExportCurrentClicked  += ExportCurrent;
            });

            Invalidate();
        }

        private bool   DestIsSrc { get; set; }

        private Models Models    { get; }

        private void Invalidate()
        => Models.ForEach(model => {
            model.DestIsSrc.IsPressed = DestIsSrc;
            model.DestIsSrc.SetImageMso(DestIsSrc.ToggleImage());
            model.ExportSelected.IsEnabled = ! DestIsSrc;
            model.DestIsSrc.IsLarge      = model.Suffix == "PG";
            model.ExportSelected.IsLarge = model.Suffix == "PG";
            model.ExportCurrent.IsLarge  = model.Suffix == "PG";

            model.Invalidate();
        });

        private void UseSrcFolderToggled(object sender, EventArgs<bool> e) {
            DestIsSrc = e.Value;

            Invalidate();
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

                ProjectFilterExcel.ExtractOpenProject(Application.ActiveWorkbook, DestIsSrc);
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
                    list[fd.FilterIndex-1].ExtractProjects(fd.SelectedItems, DestIsSrc);
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
