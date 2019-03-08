////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using Microsoft.Office.Core;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.Models;

namespace PGSolutions.BetterRibbon {
    /// <summary>Partial class interface between Designer-authored and humn-authored code.</summary>
    /// <remarks>
    /// <a href=" https://go.microsoft.com/fwlink/?LinkID=271226">Adding Ribbon XML to a project.</a>
    /// 
    /// Take care renaming this class or its namespace; and coordinate any such with the content
    /// of the (hidden) ThisAddIn.Designer.xml file. Commit frequently. Excel is very tempermental
    /// on the naming of ribbon objects and provides poor, and very minimal, diagnostic information.
    /// </remarks>
    [CLSCompliant(false)]
    public partial class ThisAddIn {
        /// <summary>The ribbon-callback dispatcher for VBA customizable ribbon tabs/groups.</summary>
        internal CustomDispatcher      Dispatcher { get; }
                = new CustomDispatcher(Properties.Resources.RibbonXml,new MyResourceManager());

        /// <summary>The VBA-accessible entry point for the ribbon dispatcher.</summary>
        private  ICustomRibbonComEntry ComEntry   => new CustomRibbonComEntry(Dispatcher);

        /// <summary>Root view-model for the VBA customizable ribbon.</summary>
        [SuppressMessage("Microsoft.Performance","CA1811:AvoidUncalledPrivateCode")]
        internal CustomRibbonViewModel ViewModel  { get; private set; }

        /// <inheritdoc/>
        protected override IRibbonExtensibility CreateRibbonExtensibilityObject() => Dispatcher;

        /// <inheritdoc/>
        protected override object RequestComAddInAutomationService() => ComEntry;

        private void ThisAddIn_Startup(object sender, EventArgs e) {
            ViewModel = new CustomRibbonViewModel(Dispatcher);

            Application.WorkbookActivate    += Dispatcher.Workbook_Activate;
            Application.WorkbookDeactivate  += Dispatcher.Workbook_Deactivate;
            Application.WorkbookBeforeSave  += Dispatcher.Workbook_BeforeSave;
            Application.WorkbookAfterSave   += Dispatcher.Workbook_AfterSave;
            Application.WorkbookBeforeClose += Dispatcher.Workbook_Close;
        }

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
