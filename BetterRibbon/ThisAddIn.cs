////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using Microsoft.Office.Core;

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
        internal CustomDispatcher      Dispatcher { get; private set; }
             //   = new CustomDispatcher(Properties.Resources.RibbonXml,new MyResourceManager());

        /// <summary>Root view-model for the VBA customizable ribbon.</summary>
        [SuppressMessage("Microsoft.Performance","CA1811:AvoidUncalledPrivateCode")]
        internal CustomRibbonViewModel ViewModel  { get; private set; }

        /// <inheritdoc/>
        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        => Dispatcher = new CustomDispatcher(Properties.Resources.RibbonXml,new MyResourceManager());

        /// <summary>Returns the VBA-accessible entry point for the ribbon dispatcher.</summary>
        protected override object RequestComAddInAutomationService() => CustomRibbonComEntry.New(Dispatcher);

        private void ThisAddIn_Startup(object sender, EventArgs e) {
            ViewModel = new CustomRibbonViewModel(Dispatcher);

            Application.WorkbookActivate    += Dispatcher.Workbook_Activate;
            Application.WorkbookDeactivate  += Dispatcher.Workbook_Deactivate;
            Application.WorkbookBeforeSave  += Dispatcher.Workbook_BeforeSave;
            Application.WorkbookAfterSave   += Dispatcher.Workbook_AfterSave;
            Application.WorkbookBeforeClose += Dispatcher.Workbook_Close;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e) { }

        private static Assembly Current_AssemblyResolve(object sender,ResolveEventArgs args) {
            const string dllName = "EmbedAssembly.PGSolutions.RibbonDispatcher.dll";

            using(var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(dllName)) {
                byte[] assemblyData = new byte[stream.Length];
                stream.Read(assemblyData, 0, assemblyData.Length);
                return Assembly.Load(assemblyData);
            }
        }

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
