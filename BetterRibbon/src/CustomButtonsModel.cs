////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Collections.Generic;
using PGSolutions.RibbonDispatcher.Utilities;

namespace PGSolutions.BetterRibbon {
    internal sealed class CustomButtonsModel {
        public CustomButtonsModel(CustomizableButtonsViewModel viewModel) =>
            ViewModel = viewModel;

        public IReadOnlyDictionary<string, IActivatable> AdaptorControls => ViewModel.AdaptorControls;
        public string             GroupId => ViewModel.GroupId;
 
        public void Invalidate() => ViewModel.Invalidate();

        private CustomizableButtonsViewModel ViewModel { get; set; }
    }
}
