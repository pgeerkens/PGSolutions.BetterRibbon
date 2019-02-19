////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.BetterRibbon {
    /// <summary>Implementation of <see cref="AbstractRibbonGroupModel"/> for the VBA-customizable ribbon controls..</summary>
    [CLSCompliant(false)]
    public sealed class CustomButtonsGroup1Model : AbstractRibbonGroupModel, IRibbonCommonSource {
        internal CustomButtonsGroup1Model(RibbonGroupViewModel viewModel) : base(viewModel) { }
    }
}
