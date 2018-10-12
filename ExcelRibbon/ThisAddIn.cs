using System;
using Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.ExcelRibbon
{
    [CLSCompliant(false)]
    public partial class ThisAddIn {
        internal RibbonViewModel ViewModel { get; private set; }

        private void ThisAddIn_Startup(object sender, EventArgs e) { }

        private void ThisAddIn_Shutdown(object sender, EventArgs e) { }

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject() 
            => ViewModel = new RibbonViewModel();

        private Lazy<IMain> ComEntry = new Lazy<IMain>(() => new Main());

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
