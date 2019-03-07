////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;

using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.BetterRibbon {

    /// <summary>The (top-level) ViewModel for the ribbon interface.</summary>
    [SuppressMessage("Microsoft.Naming","CA1710:IdentifiersShouldHaveCorrectSuffix")]
    [Description("The (top-level) ViewModel for the ribbon interface.")]
    [CLSCompliant(false)]
    public sealed class RibbonViewModel: GroupVM, IRibbonViewModel {
        internal RibbonViewModel(CustomDispatcher dispatcher) 
        : base("TabPGSolutions",dispatcher.ViewModelFactory?.ViewModelRoot)
        { }

        /// <inheritdoc/>
        internal  IGroupVM  CustomControlsGroupVM
        => Controls.Item<TabVM>(ControlId).GetControl<GroupVM>("CustomizableGroup");
    }
}
