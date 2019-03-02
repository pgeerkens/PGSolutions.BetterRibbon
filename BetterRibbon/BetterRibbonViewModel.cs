////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;

using Microsoft.Office.Core;

using PGSolutions.RibbonDispatcher.Models;
using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.BetterRibbon {

    /// <summary>The (top-level) ViewModel for the ribbon interface.</summary>
    [Description("The (top-level) ViewModel for the ribbon interface.")]
    [CLSCompliant(false)]
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
        public string    ControlId             { get; }
        public IRibbonUI RibbonUI              { get; }

        public GroupVM   BrandingGroupVM       { get; }
        public GroupVM   LinkedAnalysisGroupVM { get; }
        public GroupVM   VbaExportGroupVM_MS   { get; }
        public GroupVM   VbaExportGroupVM_PG   { get; }
        public GroupVM   CustomControlsGroupVM { get; }
    }
}
