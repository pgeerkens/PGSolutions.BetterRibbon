////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using Microsoft.Office.Core;

using PGSolutions.RibbonDispatcher.Models;
using PGSolutions.BetterRibbon.Properties;
using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.BetterRibbon {
    /// <summary>The (top-level) ViewModel for the ribbon interface.</summary>
    /// <remarks>
    /// <a href=" https://go.microsoft.com/fwlink/?LinkID=271226"> For more information about adding callback methods.</a>
    /// 
    /// Take care renaming this class, or its namespace; and coordinate any such with the content
    /// of the (hidden) ThisAddIn.Designer.xml file. Commit frequently. Excel is very tempermental
    /// on the naming of ribbon objects and provides poor, and very minimal, diagnostic information.
    /// 
    /// This class MUST be ComVisible for the ribbon to launch properly; <see cref="IRibbonExtensibility"/>.
    /// </remarks>
    [Description("The (top-level) ViewModel for the ribbon interface.")]
    [CLSCompliant(false)]
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable",
            Justification = "Public, Non-Creatable, class with exported Events.")]
    [ComVisible(true)]
    public sealed class Dispatcher: AbstractDispatcher, IRibbonExtensibility {
        internal Dispatcher(string controlId) : base(controlId, new MyResourceManager()) { }

        /// <inheritdoc/>
        protected override string RibbonXml => Resources.Ribbon;
    }

    /// <summary>The (top-level) ViewModel for the ribbon interface.</summary>
    /// <remarks>
    /// <a href=" https://go.microsoft.com/fwlink/?LinkID=271226"> For more information about adding callback methods.</a>
    /// 
    /// Take care renaming this class, or its namespace; and coordinate any such with the content
    /// of the (hidden) ThisAddIn.Designer.xml file. Commit frequently. Excel is very tempermental
    /// on the naming of ribbon objects and provides poor, and very minimal, diagnostic information.
    /// 
    /// This class MUST be ComVisible for the ribbon to launch properly; <see cref="IRibbonExtensibility"/>.
    /// </remarks>
    [Description("The (top-level) ViewModel for the ribbon interface.")]
    [CLSCompliant(false)]
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable",
            Justification = "Public, Non-Creatable, class with exported Events.")]
    [ComVisible(true)]
    public sealed class BetterRibbonViewModel: GroupVM, IRibbonViewModel {
        internal BetterRibbonViewModel(AbstractDispatcher dispatcher, string controlId)
        : base(dispatcher.ViewModelFactory,controlId) {
            ControlId             = controlId;
            RibbonUI              = dispatcher.RibbonUI;
            BrandingGroupVM       = GetControl<GroupVM>("BrandingGroup");
            LinkedAnalysisGroupVM = GetControl<GroupVM>("LinksAnalysisGroup");
            VbaExportGroupVM_MS   = GetControl<GroupVM>("VbaExportGroupMS");
            VbaExportGroupVM_PG   = GetControl<GroupVM>("VbaExportGroupPG");
            CustomControlsGroupVM = GetControl<GroupVM>("CustomizableGroup");
        }

        /// <inheritdoc/>
        public string    ControlId              { get; }
        public IRibbonUI RibbonUI               { get; }

        public GroupVM    BrandingGroupVM       { get; }
        public GroupVM    LinkedAnalysisGroupVM { get; }
        public GroupVM    VbaExportGroupVM_MS   { get; }
        public GroupVM    VbaExportGroupVM_PG   { get; }
        public GroupVM    CustomControlsGroupVM { get; }
    }
}
