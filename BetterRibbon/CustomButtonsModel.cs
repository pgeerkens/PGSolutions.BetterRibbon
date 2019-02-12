////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using PGSolutions.RibbonDispatcher.ComClasses;

namespace PGSolutions.BetterRibbon {
    internal sealed class CustomButtonsModel {
        public CustomButtonsModel(CustomizableButtonsViewModel viewModel) =>
            ViewModel = viewModel;
 
        public void   Invalidate() => ViewModel.Invalidate();

        public TControl GetControl<TControl>(string controlId) where TControl:RibbonCommon =>
            ViewModel.GetControl<TControl>(controlId);

        public void SetShowWhenInactive(bool showWhenInactive) =>
            ViewModel.SetShowWhenInactive(showWhenInactive);

        private CustomizableButtonsViewModel ViewModel { get; set; }
    }
}
