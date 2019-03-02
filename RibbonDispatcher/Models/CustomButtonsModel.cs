////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.RibbonDispatcher.Models {
    /// <summary>Implementation of <see cref="AbstractRibbonGroupModel"/> for the VBA-customizable ribbon controls..</summary>
    public sealed class CustomButtonsGroupModel : AbstractRibbonGroupModel, IControlSource {
        public CustomButtonsGroupModel(IModelFactory factory, IGroupVM viewModel)
        : base(viewModel, factory?.GetStrings(viewModel.Id)) { }
    }
}
