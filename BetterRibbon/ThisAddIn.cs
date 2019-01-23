////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

using Excel = Microsoft.Office.Interop.Excel;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.BetterRibbon {
    [CLSCompliant(false)]
    public partial class ThisAddIn {
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

        private void WorkbookDeactivate(Workbook wb) => ViewModel.DetachControls();
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
