////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using PGSolutions.RibbonDispatcher.Utilities;

namespace PGSolutions.BetterRibbon {
    internal sealed class BrandingModel {
        public BrandingModel(BrandingViewModel viewModel) =>
            ViewModel = viewModel;

        public void Attach()     => ViewModel.ButtonClicked += ButtonClicked;
        public void Detach()     => ViewModel.ButtonClicked -= ButtonClicked;
        public void Invalidate() => ViewModel.Invalidate();

        private BrandingViewModel ViewModel { get; set; }

        private void ButtonClicked(object sender) =>
            $"Canadian, eh!\n\nVersion: {Globals.ThisAddIn.VersionNo3}".MsgBoxShow();
    }
}
