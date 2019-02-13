////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using PGSolutions.RibbonDispatcher.ComClasses;

namespace PGSolutions.BetterRibbon {
    internal sealed class CustomButtonsModel {
        public CustomButtonsModel(CustomButtonsViewModel viewModel) {
            ViewModel = viewModel;

            viewModel.Attach();
        }

        public void   Invalidate() => ViewModel.Invalidate();

        public TControl GetControl<TControl>(string controlId) where TControl:RibbonCommon
        => ViewModel.GetControl<TControl>(controlId);

        private CustomButtonsViewModel ViewModel { get; set; }
    }
}
