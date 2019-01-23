////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using PGSolutions.RibbonDispatcher.ComClasses;

namespace PGSolutions.BetterRibbon {
    internal sealed class CustomButtonsModel {
        public CustomButtonsModel(CustomizableButtonsViewModel viewModel) =>
            ViewModel = viewModel;
 
        //public string GroupId      => ViewModel.GroupId;
        public void   Invalidate() => ViewModel.Invalidate();

        public TControl GetControl<TControl>(string controlId) where TControl:RibbonCommon =>
            ViewModel.GetControl<TControl>(controlId);

        public void SetShowWhenInactive(bool showWhenInactive) =>
            ViewModel.SetShowWhenInactive(showWhenInactive);

        private CustomizableButtonsViewModel ViewModel { get; set; }
    }
}
