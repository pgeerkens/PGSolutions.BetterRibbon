////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.InteropServices;
using System.Xml.Linq;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.RibbonDispatcher.Models {
    [CLSCompliant(false)]
    [ComVisible(true)]
    public abstract class AbstractCustomDispatcher: AbstractDispatcher {
        protected AbstractCustomDispatcher()
        => RibbonXmlDoc = XDocument.Parse(RibbonXml);

        private            XDocument       RibbonXmlDoc   { get; }

        private            Factories       Factories      { get; } = new Factories();

        /// <inheritdoc/>
        public override void OnRibbonLoad(IRibbonUI ribbonUI) {
            SaveCurrent(":");

            base.OnRibbonLoad(ribbonUI);
        }

        /// <inheritdoc/>
        public override void RegisterWorkbook(string workbookName) {
            if ( ! Factories.TryGetValue(workbookName,out var factory)) {
                factory = ViewModelFactory.ParseXmlDoc(RibbonXmlDoc.Root).ReKey(workbookName);
                Factories.Add(factory);
            }
            SetViewModelFactory(factory);
            System.Diagnostics.Debug.Assert(ViewModelFactory.Key == workbookName);
        }


        public void Workbook_Activate(Workbook wb) => RegisterWorkbook(wb.Name);

        public void Workbook_Deactivate(Workbook wb) { }

        public void Workbook_BeforeSave(Workbook wb, bool SaveAsUI, ref bool Cancel) => FloatCurrent();

        public void Workbook_AfterSave(Workbook wb, bool Success) => SaveCurrent(wb.Name);

        public void Workbook_Close(Workbook wb, ref bool Cancel) => FloatCurrent();

        /// <inheritdoc/>
        internal void SaveCurrent(string workbookName) {
            ViewModelFactory.ReKey(workbookName);
            if ( ! Factories.Contains(ViewModelFactory)) Factories.Add(ViewModelFactory);
            SetViewModelFactory(ViewModelFactory);
            System.Diagnostics.Debug.Assert(ViewModelFactory.Key == workbookName);
        }

        /// <inheritdoc/>
        internal void FloatCurrent() {
            if (Factories.Contains(ViewModelFactory)) Factories.Remove(ViewModelFactory);
        }
    }
}
