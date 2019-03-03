////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Linq;
using PGSolutions.RibbonDispatcher;
using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.Models;
using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.BetterRibbon {
    /// <summary>The (top-level) TabModel for the ribbon interface.</summary>
    [CLSCompliant(false)]
    public sealed class RibbonModel : AbstractRibbonTabModel {
        internal RibbonModel(RibbonViewModel viewModel, Func<string,IControlStrings> func)
        : base(viewModel, new List<ICanInvalidate> {
                new CustomButtonsGroupModel(func, viewModel.CustomControlsGroupVM)
            }.AsReadOnly())
        => CustomGroupModel = Models.OfType<CustomButtonsGroupModel>().FirstOrDefault();

        internal CustomButtonsGroupModel CustomGroupModel { get; }
    }
}
