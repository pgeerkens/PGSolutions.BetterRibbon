////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////

using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    /// <summary>Implementation of <see cref="AbstractRibbonGroupModel"/> for the VBA-customizable ribbon controls..</summary>
    public sealed class CustomButtonsGroupModel : AbstractRibbonGroupModel, IControlSource {
        public CustomButtonsGroupModel(IRibbonViewModel viewModel, AbstractModelFactory factory, string viewModelName)
        : base(viewModel, viewModelName, factory.GetStrings(viewModelName)) { }
    }
}
