////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017-8 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

using PGSolutions.RibbonUtilities;
using PGSolutions.RibbonUtilities.LinksAnalysis;
using PGSolutions.RibbonUtilities.VbaSourceExport;

namespace PGSolutions.BetterRibbon {
    using static RibbonDispatcher.ComClasses.Extensions;
    using Models        = IReadOnlyList<VbaSourceExportGroupModel>;
    using ComInterfaces = RibbonDispatcher.ComInterfaces;

    /// <summary>The TabModel for the VBA Source Export Group on the BetterRibbon.</summary>
    internal sealed class VbaSourceExportModel : ComInterfaces.IInvalidate {
        /// <summary>.</summary>
        /// <param name="models"></param>
        public VbaSourceExportModel(Models models) {
            Models    = models;
            DestIsSrc = false;

            foreach (var model in Models) {
                model.UseSrcFolderToggled   += UseSrcFolderToggled;
                model.ExportSelectedClicked += ExportSelected;
                model.ExportCurrentClicked  += ExportCurrent;
            }

            Invalidate();
        }

        private bool   DestIsSrc { get; set; }

        private Models Models    { get; }

        public void Invalidate() {
            foreach (var model in Models) {
                model.DestIsSrc.IsPressed = DestIsSrc;
                model.DestIsSrc.SetImageMso(DestIsSrc.ToggleImage());
                model.ExportSelected.IsEnabled = ! DestIsSrc;
                model.DestIsSrc.IsLarge      = model.Suffix == "PG";
                model.ExportSelected.IsLarge = model.Suffix == "PG";
                model.ExportCurrent.IsLarge  = model.Suffix == "PG";

                model.Invalidate();
        }   }

        private void UseSrcFolderToggled(object sender, ComInterfaces.EventArgs<bool> e) {
            DestIsSrc = e.Value;

            Invalidate();
        }

        /// <summary>Writes a status message to the Excel Application Status Bar.</summary>
        private void StatusAvailable(object sender, EventArgs<string> e)
        => Application.StatusBar = e.Value;

        /// <summary>Extracts VBA modules from current EXCEL workbook to a sibling directory.</summary>
        /// <param name="sender">The object that initiated the event.</param>
        /// <remarks>
        /// Requires that access to the VBA project object model be trusted (Macro Security).
        /// </remarks>
        private void ExportCurrent(object sender, EventArgs e) {
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
        [SuppressMessage("Microsoft.Reliability", "CA2001:AvoidCallingProblematicMethods", MessageId = "System.GC.Collect")]
        private void ExportSelected(object sender, EventArgs e) {
            if (!IsProjectModelTrusted()) { return; }

            var fd = Application.FileDialog[MsoFileDialogType.msoFileDialogFilePicker];
            fd.Title = "Select VBA Project(s) to Export From";
            fd.ButtonName = "Export";
            fd.AllowMultiSelect = true;
            fd.Filters.Clear();
            fd.InitialFileName = Application.ActiveWorkbook?.Path ?? "C:\\";

            Application.Cursor = XlMousePointer.xlWait;
            StatusAvailable(this, new EventArgs<string>("Loading background processor ..."));
            using (var processor = WorkbookProcessor.New(Application, true)) {
                var list = VbaSourceExporter.FillFilters(processor, fd);
                Application.Cursor = XlMousePointer.xlDefault;
                if (fd.Show() != 0) {
                    Application.Cursor = XlMousePointer.xlWait;
                    try {
                        var exporter = new VbaSourceExporter(Application);
                        exporter.StatusAvailable += StatusAvailable;
                        exporter.ExportSelected(list[fd.FilterIndex-1], fd.SelectedItems, DestIsSrc);
                        exporter.StatusAvailable -= StatusAvailable;
                    }
                    catch (IOException ex) { ex.Message.MsgBoxShow(CallerName()); }
                    finally {
                        Application.Cursor = XlMousePointer.xlDefault;
                    }
                }
            }
        #if DEBUG
            Application.Cursor = XlMousePointer.xlWait;
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Application.Cursor = XlMousePointer.xlDefault;
        #endif
            StatusAvailable(this, new EventArgs<string>("Ready"));
        }

        private static bool IsProjectModelTrusted() {
            try { return Application.VBE != null; }
            catch (COMException) { PleaseEnableTrust(); }
            catch (InvalidOperationException) { PleaseEnableTrust(); }
            return false;
        }

        private static void PleaseEnableTrust()
        => new StringBuilder()
            .AppendLine("VBA Export requires trust of the VBA Project object model.")
            .AppendLine()
            .AppendLine("Please enable trust at:")
            .AppendLine("    File")
            .AppendLine("        -> Options")
            .AppendLine("        -> Trust Center")
            .AppendLine("        -> Trust Center Settings")
            .AppendLine("        -> Macro Settings")
            .AppendLine("        -> Trust Access to the VBA Project object model")
            .AppendLine()
            .AppendLine(" or:")
            .AppendLine("    Developer")
            .AppendLine("        -> Macro Security")
            .AppendLine("        -> Trust Access to the VBA Project object model")
            .ToString().MsgBoxShow("VBA Project Object Model Not Trusted");

        private static Application Application => Globals.ThisAddIn.Application;
    }
}
