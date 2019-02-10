////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonDispatcher.Utilities;
using PGSolutions.RibbonUtilities.VbaSourceExport;
using System;

namespace PGSolutions.BetterRibbon {
    internal sealed class BrandingModel {
        public BrandingModel(BrandingViewModel viewModel) =>
            ViewModel = viewModel;

        public void Attach()     => ViewModel.ButtonClicked += ButtonClicked;
        public void Detach()     => ViewModel.ButtonClicked -= ButtonClicked;
        public void Invalidate() => ViewModel.Invalidate();

        private BrandingViewModel ViewModel { get; set; }

        private void ButtonClicked(object sender) =>
            ( $"PGSolutions Better Ribbon\n\n"
            + $"Better Ribbon V {Globals.ThisAddIn.VersionNo3}\n"
            + $"Ribbon Utilities V {UtilitiesVersion.Format2()}\n"
            + $"Ribbon Dispatcher V {DispatcherVersion.Format2()}\n\n"
            + $"{ViewModel.BrandingButton.SuperTip}"
            ).MsgBoxShow();

        Version DispatcherVersion => new RibbonFactory().GetType().Assembly.GetName().Version;
        Version UtilitiesVersion  => new VbaExportEventArgs(null).GetType().Assembly.GetName().Version;
    }
}
