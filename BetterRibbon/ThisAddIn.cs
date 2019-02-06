////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Deployment.Application;
using System.Diagnostics.CodeAnalysis;

using Microsoft.Office.Core;

using Excel = Microsoft.Office.Interop.Excel;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.Utilities;

namespace PGSolutions.BetterRibbon {
    [CLSCompliant(false)]
    public partial class ThisAddIn {
        public string VersionNo => ApplicationDeployment.IsNetworkDeployed
            ? ApplicationDeployment.CurrentDeployment.CurrentVersion.FormatVersion()
            : null;
        public string VersionNo2 => System.Windows.Forms.Application.ProductVersion;
        public string VersionNo3 => System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.FormatVersion();

        internal BetterRibbonViewModel ViewModel { get; private set; }

        private void ThisAddIn_Startup(object sender, EventArgs e) {
            //Application.WorkbookDeactivate += WorkbookDeactivate;
            Application.WindowDeactivate += WindowDeactivate;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e) { }

        /// <summary>.</summary>
        protected override IRibbonExtensibility CreateRibbonExtensibilityObject() 
            => ViewModel = new BetterRibbonViewModel();

        private Lazy<Main> ComEntry = new Lazy<Main>(() => new Main());

        /// <summary>.</summary>
        protected override object RequestComAddInAutomationService() =>
            ComEntry.Value as IBetterRibbon;

        [SuppressMessage( "Microsoft.Usage", "CA1801:ReviewUnusedParameters", MessageId = "wb" )]
        private void WorkbookDeactivate(Workbook wb) => ViewModel.DetachControls();
        [SuppressMessage( "Microsoft.Usage", "CA1801:ReviewUnusedParameters", MessageId = "wb" )]
        [SuppressMessage( "Microsoft.Usage", "CA1801:ReviewUnusedParameters", MessageId = "wn" )]
        private void WindowDeactivate(Workbook wb, Excel.Window wn) => ViewModel.DetachControls();

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
