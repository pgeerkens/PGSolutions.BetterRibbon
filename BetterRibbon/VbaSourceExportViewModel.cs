////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017-8 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonDispatcher.Utilities;
using PGSolutions.RibbonUtilities.VbaSourceExport;

namespace PGSolutions.BetterRibbon {

    /// <summary>.</summary>
    [CLSCompliant(false)]
    public sealed class VbaSourceExportViewModel : AbstractVbaSourceExportViewModel, IVbaSourceExportViewModel, IApplication {
        public VbaSourceExportViewModel(IRibbonFactory factory, string suffix) : base(factory, suffix) { }

        protected override void OnExportCurrent(object sender) {
            if (!IsProjectModelTrusted()) { return; }
            var securitySaved = Application.AutomationSecurity;
            Application.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;

            try { base.OnExportCurrent(sender); }
            catch (IOException ex) { ex.Message.MsgBoxShow("OnExportCurrent"); }
            finally { Application.AutomationSecurity = securitySaved; }
        }

        protected override void OnExportSelected(object sender) {
            if (!IsProjectModelTrusted()) { return; }
            var securitySaved = Application.AutomationSecurity;
            Application.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;

            try { base.OnExportSelected(sender); }
            catch (IOException ex) { ex.Message.MsgBoxShow("OnExportSelected"); }
            finally { Application.AutomationSecurity = securitySaved; }
        }

        private bool IsProjectModelTrusted() {
            try { return Application.VBE != null; }
            catch (COMException) { PleaseEnableTrust(); }
            catch (InvalidOperationException) { PleaseEnableTrust(); }
            return false;
        }

        private static void PleaseEnableTrust()
        => "Please enable trust of the Project Object Model".MsgBoxShow("Project Model Not Trusted");

        private static string ToggleImage(bool isPressed) => isPressed ? "TagMarkComplete" : "MarginsShowHide";

        protected override Application Application => Globals.ThisAddIn.Application;

        /// <inheritdoc/>
        protected override Workbook ActiveWorkbook => Application.ActiveWorkbook;

        /// <inheritdoc/>
        public override bool DisplayAlerts {
            get => Application.DisplayAlerts;
            set => Application.DisplayAlerts = value;
        }

        /// <inheritdoc/>
        public override dynamic StatusBar {
            get => Application.StatusBar;
            set => Application.StatusBar = value;
        }

        /// <inheritdoc/>
        public override MsoAutomationSecurity AutomationSecurity {
            get => Application.AutomationSecurity;
            protected set => Application.AutomationSecurity = value;
        }
    }
}
