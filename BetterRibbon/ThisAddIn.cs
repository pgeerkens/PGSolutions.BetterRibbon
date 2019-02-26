////////////////////////////////////////////////////////////////////////////////////////////////////
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
        protected override IRibbonExtensibility CreateRibbonExtensibilityObject() {
            ViewModel = new BetterRibbonViewModel("TabPGSolutions");
            ViewModel.Initialized += ViewModel_Initialized;
            return ViewModel;
        }

        private void ViewModel_Initialized(object sender, EventArgs e) {
            Model = new BetterRibbonModel(ViewModel, 
                    ViewModel.ViewModelFactory.NewModelFactory2(new MyResourceManager()));
            ViewModel.Initialized -= ViewModel_Initialized;
        }

        private void ThisAddIn_Startup(object sender, EventArgs e) {
            Application.WorkbookDeactivate += Workbook_Deactivate;
            Application.WindowDeactivate += Window_Deactivate;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e) { }

        /// <summary>.</summary>
        protected override object RequestComAddInAutomationService()
        => ComEntry as IBetterRibbon;

        internal BetterRibbonViewModel ViewModel { get; private set; }

        internal BetterRibbonModel     Model     { get; private set; }

        private  Main                  ComEntry  => new Main(ViewModel.ViewModelFactory.NewModelFactory);

        /// <summary>.</summary>
        public static string VersionNo => ApplicationDeployment.IsNetworkDeployed
            ? ApplicationDeployment.CurrentDeployment.CurrentVersion?.Format()
            : null;

        /// <summary>.</summary>
        public static string VersionNo2 => System.Windows.Forms.Application.ProductVersion;

        /// <summary>.</summary>
        public string VersionNo3 =>GetType().Assembly.GetName().Version?.Format();

        [SuppressMessage( "Microsoft.Usage", "CA1801:ReviewUnusedParameters", MessageId = "wb" )]
        private void Workbook_Deactivate(Workbook wb)
        => Model?.DetachCustomControls();

        [SuppressMessage( "Microsoft.Usage", "CA1801:ReviewUnusedParameters", MessageId = "wb" )]
        [SuppressMessage( "Microsoft.Usage", "CA1801:ReviewUnusedParameters", MessageId = "wn" )]
        private void Window_Deactivate(Workbook wb, Excel.Window wn) 
        => Model?.DetachCustomControls();

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
