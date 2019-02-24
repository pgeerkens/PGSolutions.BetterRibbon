////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ComClasses.ViewModels;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    /// <summary>Implementation of <see cref="AbstractRibbonGroupModel"/> for the VBA-customizable ribbon controls..</summary>
    [CLSCompliant(false)]
    public sealed class CustomButtonsGroupModel : AbstractRibbonGroupModel, IControlSource {
        public CustomButtonsGroupModel(IRibbonViewModel viewModel, string viewModelName)
        : base(viewModel, viewModelName) { }
    }
}
