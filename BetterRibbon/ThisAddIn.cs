////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Deployment.Application;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

using PGSolutions.RibbonDispatcher;
using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.BetterRibbon {

    /// <summary>Partial class interface between Designer-authored and humn-authored code.</summary>
    /// <remarks>
    /// <a href=" https://go.microsoft.com/fwlink/?LinkID=271226"> For more information about adding callback methods.</a>
    /// 
    /// Take care renaming this class, or its namespace; and coordinate any such with the content
    /// of the (hidden) ThisAddIn.Designer.xml file. Commit frequently. Excel is very tempermental
    /// on the naming of ribbon objects and provides poor, and very minimal, diagnostic information.
    /// </remarks>
    [CLSCompliant(false)]
    public partial class ThisAddIn {
        private void ThisAddIn_Startup(object sender, EventArgs e) {
            Dispatcher.RegisterWorkbook(":");
            ViewModel = new RibbonViewModel(Dispatcher);

            Application.WorkbookActivate    += Dispatcher.Workbook_Activate;
            Application.WorkbookDeactivate  += Dispatcher.Workbook_Deactivate;
            Application.WorkbookBeforeSave  += Dispatcher.Workbook_BeforeSave;
            Application.WorkbookAfterSave   += Dispatcher.Workbook_AfterSave;
            Application.WorkbookBeforeClose += Dispatcher.Workbook_Close;
        }

        internal CustomDispatcher      Dispatcher { get; } = new CustomDispatcher();
        internal RibbonViewModel       ViewModel  { get; private set; }
        private  ICustomRibbonComEntry ComEntry   => new ComEntry(Dispatcher);

        /// <inheritdoc/>
        protected override IRibbonExtensibility CreateRibbonExtensibilityObject() => Dispatcher;

        /// <inheritdoc/>
        protected override object RequestComAddInAutomationService() => ComEntry;

        internal void RegisterWorkbook(string workbookName) => Dispatcher.RegisterWorkbook(workbookName);

        private void ThisAddIn_Shutdown(object sender, EventArgs e) { }

        /// <summary>.</summary>
        public static string VersionNo => ApplicationDeployment.IsNetworkDeployed
            ? ApplicationDeployment.CurrentDeployment.CurrentVersion?.Format()
            : new Version(0,0,0,0).Format();

        /// <summary>.</summary>
        public static string VersionNo2 => System.Windows.Forms.Application.ProductVersion;

        /// <summary>.</summary>
        public static string VersionNo3 => typeof(ThisAddIn).Assembly.GetName().Version?.Format();

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
