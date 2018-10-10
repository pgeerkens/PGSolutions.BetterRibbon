using System;
using Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;

using PGSolutions.RibbonDispatcher;

namespace PGSolutions.ExcelRibbon {
    [CLSCompliant(false)]
    public partial class ThisAddIn {
        private RibbonViewModel _viewModel;

        private void ThisAddIn_Startup(object sender, EventArgs e) { }

        private void ThisAddIn_Shutdown(object sender, EventArgs e) { }

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject() 
            => _viewModel = new RibbonViewModel();

        private Lazy<Main> ComEntry = new Lazy<Main>(() => new Main());

        protected override object RequestComAddInAutomationService() => ComEntry.Value;

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
