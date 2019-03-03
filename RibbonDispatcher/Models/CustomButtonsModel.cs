////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.RibbonDispatcher.Models {
    /// <summary>Implementation of <see cref="AbstractRibbonGroupModel"/> for the VBA-customizable ribbon controls..</summary>
    public sealed class CustomButtonsGroupModel : AbstractRibbonGroupModel, IControlSource {
        public CustomButtonsGroupModel( Func<string,IControlStrings> func, IGroupVM viewModel)
        : base(viewModel, func(viewModel.ControlId)) { }
    }
}
