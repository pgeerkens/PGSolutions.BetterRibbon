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
        : base("TabPGSolutions",dispatcher.ViewModelFactory?.ViewModelRoot)
        => RibbonUI = dispatcher.RibbonUI;
                
        /// <inheritdoc/>
        public  IRibbonUI RibbonUI { get; }

        public  IGroupVM  CustomControlsGroupVM
        => Controls.Item<TabVM>(ControlId).GetControl<GroupVM>("CustomizableGroup");
    }
}
