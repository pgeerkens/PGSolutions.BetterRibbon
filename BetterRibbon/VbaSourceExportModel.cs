////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017-8 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

using PGSolutions.RibbonUtilities;
using PGSolutions.RibbonUtilities.VbaSourceExport;

namespace PGSolutions.BetterRibbon {
    using static RibbonDispatcher.ComClasses.Extensions;
    using Models        = List<VbaSourceExportGroupModel>;
    using ComInterfaces = RibbonDispatcher.ComInterfaces;
    using System.Text;

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

        public void Invalidate()
        => Models.ForEach(model => {
            model.DestIsSrc.IsPressed = DestIsSrc;
            model.DestIsSrc.SetImageMso(DestIsSrc.ToggleImage());
            model.ExportSelected.IsEnabled = ! DestIsSrc;
            model.DestIsSrc.IsLarge      = model.Suffix == "PG";
            model.ExportSelected.IsLarge = model.Suffix == "PG";
            model.ExportCurrent.IsLarge  = model.Suffix == "PG";

            model.Invalidate();
        });

        private void UseSrcFolderToggled(object sender, ComInterfaces.EventArgs<bool> e) {
            DestIsSrc = e.Value;

            Invalidate();
        }

        private void StatusAvailable(object sender, EventArgs<string> e)
        => Application.StatusBar = e.Value;

        /// <summary>Extracts VBA modules from current EXCEL workbook to a sibling directory.</summary>
        /// <param name="sender">The object that initiated the event.</param>
        /// <remarks>
        /// Requires that access to the VBA project object model be trusted (Macro Security).
        /// </remarks>
        private void ExportCurrent(object sender) {
            if (!IsProjectModelTrusted()) { return; }

            var exporter = new VbaSourceExporter(Application);
            exporter.StatusAvailable += StatusAvailable;
            try {
                Application.Cursor = XlMousePointer.xlWait;
                exporter.ExtractOpenProject(Application.ActiveWorkbook, DestIsSrc);
            }
            catch (IOException ex) { ex.Message.MsgBoxShow(CallerName()); }
            finally {
                Application.Cursor = XlMousePointer.xlDefault;

                exporter.StatusAvailable -= StatusAvailable;
                Application.StatusBar = "Ready";
            }
        }

        /// <summary>Extracts VBA modules from a selected EXCEL workbook to a sibling directory.</summary>
        /// <param name="sender">The object that initiated the event.</param>
        /// <remarks>
        /// Requires that access to the VBA project object model be trusted (Macro Security).
        /// </remarks>
        private void ExportSelected(object sender) {
            if (!IsProjectModelTrusted()) { return; }

            var fd = Application.FileDialog[MsoFileDialogType.msoFileDialogFilePicker];
            fd.Title = "Select VBA Project(s) to Export From";
            fd.ButtonName = "Export";
            fd.AllowMultiSelect = true;
            fd.Filters.Clear();
            fd.InitialFileName = Application.ActiveWorkbook?.Path ?? "C:\\";

            var exporter = new VbaSourceExporter(Application);
            var list = exporter.FillFilters(fd);
            if (fd.Show() != 0) {
                try {
                    exporter.StatusAvailable += StatusAvailable;
                    exporter.ExportSelected(list[fd.FilterIndex-1], fd.SelectedItems, DestIsSrc);
                    exporter.StatusAvailable -= StatusAvailable;
                }
                catch (IOException ex) { ex.Message.MsgBoxShow(CallerName()); }
            }
        }

        private static bool IsProjectModelTrusted() {
            try { return Application.VBE != null; }
            catch (COMException) { PleaseEnableTrust(); }
            catch (InvalidOperationException) { PleaseEnableTrust(); }
            return false;
        }

        private static void PleaseEnableTrust()
        => new StringBuilder()
            .AppendLine("Please enable trust of the Project object model:")
            .AppendLine("    File -> Options")
            .AppendLine("         -> Trust Center")
            .AppendLine("         -> Trust Center Settings")
            .AppendLine("         -> Macro Settings")
            .AppendLine("         -> Trust Access to the VBA Project object model")
            .ToString().MsgBoxShow("Project Model Not Trusted");

        private static Application Application => Globals.ThisAddIn.Application;
    }
}
