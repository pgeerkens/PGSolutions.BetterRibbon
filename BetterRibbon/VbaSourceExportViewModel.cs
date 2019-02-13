////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017-8 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.IO;
using System.Runtime.InteropServices;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonUtilities.VbaSourceExport;

namespace PGSolutions.BetterRibbon {
    using static RibbonDispatcher.Utilities.Extensions;

    /// <summary>.</summary>
    [CLSCompliant(false)]
    public sealed class VbaSourceExportViewModel : AbstractVbaSourceExportViewModel, IVbaSourceExportViewModel {
        public VbaSourceExportViewModel(IRibbonFactory factory, string suffix, bool isVisible = true, bool isEnabled = true)
        : base(factory, suffix, "VbaExportGroup", isVisible, isEnabled) { }

        public override void ExportCurrent(object sender) {
            if (!IsProjectModelTrusted()) { return; }
            var securitySaved = Application.AutomationSecurity;
            Application.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;

            try {
                Application.Cursor = XlMousePointer.xlWait;
                Application.StatusBar = "Exporting VBA Source ...";

                OnExportCurrent(sender,
                        new VbaExportCurrentEventArgs(new ProjectFilterExcel(this), ActiveWorkbook));
            }
            catch (IOException ex) { ex.Message.MsgBoxShow(CallerName()); }
            finally {
                Application.AutomationSecurity = securitySaved;
                Application.StatusBar = "Ready";

                Application.Cursor = XlMousePointer.xlDefault;
            }
        }

        public override void ExportSelected(object sender) {
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

                var list = new ProjectFilters(this);
                foreach (var item in list) {
                    fd.Filters.Add(item.Description, item.Extensions);
                }
                if (fd.Show() != 0) {
                    OnExportSelected(sender,
                            new VbaExportSelectedEventArgs(list[fd.FilterIndex-1], fd.SelectedItems));
                }
            }
            catch (IOException ex) { ex.Message.MsgBoxShow(CallerName()); }
            finally {
                Application.AutomationSecurity = securitySaved;
            }
        }

        private bool IsProjectModelTrusted() {
            try { return Application.VBE != null; }
            catch (COMException) { PleaseEnableTrust(); }
            catch (InvalidOperationException) { PleaseEnableTrust(); }
            return false;
        }

        private static void PleaseEnableTrust()
        => "Please enable trust of the Project Object Model".MsgBoxShow("Project Model Not Trusted");

        protected override Application Application => Globals.ThisAddIn.Application;

        /// <inheritdoc/>
        protected override Workbook ActiveWorkbook => Application.ActiveWorkbook;

        /// <inheritdoc/>
        public override bool DisplayAlerts {
            get => Application.DisplayAlerts;
            set => Application.DisplayAlerts = value;
        }
    }
}
