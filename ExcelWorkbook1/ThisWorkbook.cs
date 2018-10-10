////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Data;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Core;
using PGSolutions.RibbonDispatcher.ConcreteCOM;

namespace PGSolutions.SampleRibbon {
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("ExampleRibbon")]
    public partial class ThisWorkbook {
        private void ThisWorkbook_Startup(object sender, EventArgs e) { }

        private void ThisWorkbook_Shutdown(object sender, EventArgs e) { }

        private Lazy<RibbonViewModel> _viewModel = new Lazy<RibbonViewModel>(() => new RibbonViewModel());

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject() => _viewModel.Value;

        protected override object GetAutomationObject() => _viewModel.Value;

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisWorkbook_Startup);
            this.Shutdown += new System.EventHandler(ThisWorkbook_Shutdown);
        }

        #endregion

    }
}
