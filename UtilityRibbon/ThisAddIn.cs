////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace PGSolutions.ToolsRibbon {
    /// <summary>Partial class interface between Designer-authored and humn-authored code.</summary>
    /// <remarks>
    /// <a href=" https://go.microsoft.com/fwlink/?LinkID=271226"> For more information about adding callback methods.</a>
    /// 
    /// Take care renaming this class, or its namespace; and coordinate any such with the content
    /// of the (hidden) ThisAddIn.Designer.xml file. Commit frequently. Excel is very tempermental
    /// on the naming of ribbon objects and provides poor, and very minimal, diagnostic information.
    /// </remarks>
    [CLSCompliant(true)]
    [ProgId("PGSolutions.ToolsRibbon")]
    public partial class ThisAddIn {
        private Dispatcher      Dispatcher { get; } = new Dispatcher();

        private IToolsComEntry  ComEntry   { get; } = new ToolsComEntry();

        private RibbonViewModel ViewModel  { get; set; }

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject() => Dispatcher;

        protected override object RequestComAddInAutomationService() => ComEntry;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        => ViewModel = new RibbonViewModel(Dispatcher);

        private void ThisAddIn_Shutdown(object sender, EventArgs e) { }

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
