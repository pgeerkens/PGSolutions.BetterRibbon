﻿////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Deployment.Application;
using System.Diagnostics.CodeAnalysis;

using Microsoft.Office.Core;

using PGSolutions.RibbonDispatcher.ComClasses;

namespace PGSolutions.BetterRibbon {
    using Excel    = Microsoft.Office.Interop.Excel;
    using Workbook = Microsoft.Office.Interop.Excel.Workbook;

    [CLSCompliant(false)]
    public partial class ThisAddIn {
        /// <summary>.</summary>
        public static string VersionNo => ApplicationDeployment.IsNetworkDeployed
            ? ApplicationDeployment.CurrentDeployment.CurrentVersion?.Format()
            : null;
        /// <summary>.</summary>
        public static string VersionNo2 => System.Windows.Forms.Application.ProductVersion;
        /// <summary>.</summary>
        public string VersionNo3 =>GetType().Assembly.GetName().Version?.Format();

        internal BetterRibbonViewModel ViewModel { get; private set; }

        private void ThisAddIn_Startup(object sender, EventArgs e) {
            Application.WorkbookDeactivate += WorkbookDeactivate;
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
