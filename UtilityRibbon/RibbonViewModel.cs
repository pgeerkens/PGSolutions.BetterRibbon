////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;

using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.ToolsRibbon {
    /// <summary>The (top-level) RibbonViewModel for the ribbon interface.</summary>
    [Description("The (top-level) RibbonViewModel for the ribbon interface.")]
    [CLSCompliant(false)]
    internal sealed class RibbonViewModel: GroupVM, IRibbonViewModel {
        public RibbonViewModel(Dispatcher dispatcher) 
        : base("TabPGSolutions",dispatcher.ViewModelFactory?.ViewModelRoot)
        { }
                
        /// <inheritdoc/>
        private TabVM TabMS => Controls.Item<TabVM>("TabDeveloper");
        private TabVM TabPG => Controls.Item<TabVM>(ControlId);

        public  IGroupVM  BrandingGroupVM       => TabPG.GetControl<GroupVM>("BrandingGroup");
        public  IGroupVM  LinkedAnalysisGroupVM => TabPG.GetControl<GroupVM>("LinksAnalysisGroup");
        public  IGroupVM  VbaExportGroupVM_MS   => TabMS.GetControl<GroupVM>("VbaExportGroupMS");
        public  IGroupVM  VbaExportGroupVM_PG   => TabPG.GetControl<GroupVM>("VbaExportGroupPG");
    }
}
