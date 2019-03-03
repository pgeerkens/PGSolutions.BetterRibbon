////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Linq;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.Models;
using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.BetterRibbon {
    /// <summary>The (top-level) TabModel for the ribbon interface.</summary>
    [CLSCompliant(false)]
    public sealed class BetterRibbonModel : AbstractRibbonTabModel {
        internal BetterRibbonModel(BetterRibbonViewModel viewModel, IModelFactory factory)
        : base(viewModel, new List<ICanInvalidate> {
                new CustomButtonsGroupModel(factory, viewModel.CustomControlsGroupVM)
            }.AsReadOnly())
        => CustomGroupModel = Models.OfType<CustomButtonsGroupModel>().FirstOrDefault();

        internal CustomButtonsGroupModel CustomGroupModel { get; }
    }
}
