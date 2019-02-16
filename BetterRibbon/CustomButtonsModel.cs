////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.BetterRibbon {
    internal sealed class CustomButtonsGroupModel : AbstractRibbonGroupModel, IRibbonCommonSource {
        public CustomButtonsGroupModel(RibbonGroupViewModel viewModel) : base(viewModel,null) { }
    }
}
