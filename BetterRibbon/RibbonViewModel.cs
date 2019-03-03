////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;

using Microsoft.Office.Core;

using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.BetterRibbon {

    /// <summary>The (top-level) ViewModel for the ribbon interface.</summary>
    [Description("The (top-level) ViewModel for the ribbon interface.")]
    [CLSCompliant(false)]
    public sealed class RibbonViewModel: GroupVM, IRibbonViewModel {
        internal RibbonViewModel(Dispatcher dispatcher) 
        : base("pg:TabPGSolutions",dispatcher.ViewModelFactory?.TabViewModels)
        => RibbonUI      = dispatcher.RibbonUI;
                
        /// <inheritdoc/>
        public  IRibbonUI RibbonUI { get; }
        private TabVM     TabMS    => Controls.Item<TabVM>("TabDeveloper");
        private TabVM     TabPG    => Controls.Item<TabVM>(ControlId);

        public  IGroupVM  BrandingGroupVM       => TabPG.GetControl<GroupVM>("BrandingGroup");
        public  IGroupVM  LinkedAnalysisGroupVM => TabPG.GetControl<GroupVM>("LinksAnalysisGroup");
        public  IGroupVM  VbaExportGroupVM_MS   => TabMS.GetControl<GroupVM>("VbaExportGroupMS");
        public  IGroupVM  VbaExportGroupVM_PG   => TabPG.GetControl<GroupVM>("VbaExportGroupPG");
        public  IGroupVM  CustomControlsGroupVM => TabPG.GetControl<GroupVM>("CustomizableGroup");
    }
}
