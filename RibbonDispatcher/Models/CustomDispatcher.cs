////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using System.Xml.Linq;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.RibbonDispatcher.Models {
    /// <summary>The concrete instantiation of <see cref="AbstractDispatcher"/> for <see cref="ThisAddIn"/>.</summary>
    /// <remarks>
    /// 
    /// This class MUST be ComVisible for the ribbon to launch properly;
    /// <see cref="IRibbonExtensibility"/>.
    /// 
    /// </remarks>
    [Description("The (top-level) ViewModel for the ribbon interface - MUST be COM-visible")]
    [CLSCompliant(true)]
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable",
            Justification = "Public, Non-Creatable, class with exported methods.")]
    [ComVisible(true)]
    public class CustomDispatcher: AbstractDispatcher, IRibbonExtensibility {
        [SuppressMessage("Microsoft.Usage","CA2214:DoNotCallOverridableMethodsInConstructors")]
        public CustomDispatcher(string ribbonXml, IResourceLoader loader){
            ResourceLoader = loader;
            RibbonXml      = ribbonXml;
            RibbonXDoc     = XDocument.Parse(RibbonXml);
        }
        /// <inheritdoc/>
        public    override IResourceLoader ResourceLoader { get; }

        /// <inheritdoc/>
        protected override string    RibbonXml  { get; }

        private            XDocument RibbonXDoc { get; }

        private            Factories Factories  { get; } = new Factories();

        /// <inheritdoc/>
        public override void OnRibbonLoad(IRibbonUI ribbonUI) {
            SaveCurrent(InvalidFileName);

            base.OnRibbonLoad(ribbonUI);
        }

        /// <inheritdoc/>
        public override void RegisterWorkbook(string workbookName) {
            if (workbookName == null) throw new ArgumentNullException(nameof(workbookName));

            if ( ! Factories.TryGetValue(workbookName,out var factory)) {
                factory = ViewModelFactory.ParseXmlDoc(RibbonXDoc.Root).Rekey(workbookName);
                Factories.Add(factory);
            }
            SetViewModelFactory(factory);
            System.Diagnostics.Debug.Assert(ViewModelFactory.Key == workbookName);
        }


        [SuppressMessage("Microsoft.Naming","CA1707:IdentifiersShouldNotContainUnderscores")]
        public void Workbook_Activate(Workbook wb) => RegisterWorkbook(wb?.Name);

        [SuppressMessage("Microsoft.Naming","CA1707:IdentifiersShouldNotContainUnderscores")]
        public void Workbook_Deactivate(Workbook wb) => RegisterWorkbook(InvalidFileName);

        [SuppressMessage("Microsoft.Naming","CA1707:IdentifiersShouldNotContainUnderscores")]
        public void Workbook_BeforeSave(Workbook wb, bool SaveAsUI, ref bool Cancel) => FloatCurrent();

        [SuppressMessage("Microsoft.Naming","CA1707:IdentifiersShouldNotContainUnderscores")]
        public void Workbook_AfterSave(Workbook wb, bool Success) => SaveCurrent(wb?.Name);

        [SuppressMessage("Microsoft.Naming","CA1707:IdentifiersShouldNotContainUnderscores")]
        public void Workbook_Close(Workbook wb, ref bool Cancel) => FloatCurrent();

        /// <inheritdoc/>
        internal void SaveCurrent(string workbookName) {
            if (workbookName == null) throw new ArgumentNullException(nameof(workbookName));

            ViewModelFactory.Rekey(workbookName);
            if ( ! Factories.Contains(ViewModelFactory)) Factories.Add(ViewModelFactory);
            SetViewModelFactory(ViewModelFactory);
            System.Diagnostics.Debug.Assert(ViewModelFactory.Key == workbookName);
        }

        /// <inheritdoc/>
        internal void FloatCurrent() {
            if (Factories.Contains(ViewModelFactory)) Factories.Remove(ViewModelFactory);
        }

        /// <summary>":" is an invalid file name - so can never be the name of a workbook</summary>
        private const string InvalidFileName = ":";
    }
}
