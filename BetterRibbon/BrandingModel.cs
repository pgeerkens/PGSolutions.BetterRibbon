////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Text;

using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonDispatcher.Utilities;
using PGSolutions.RibbonUtilities.VbaSourceExport;

namespace PGSolutions.BetterRibbon {
    internal sealed class BrandingModel {
        public BrandingModel(RibbonGroupViewModel viewModel) {
            ViewModel = viewModel;
            ViewModel.GetControl<RibbonButton>("BrandingButton").Clicked += ButtonClicked;

            ViewModel.Attach();
        }

        public void Invalidate() => ViewModel.Invalidate();

        private RibbonGroupViewModel ViewModel { get; }
        private RibbonButton BrandingButton => ViewModel.GetControl<RibbonButton>("BrandingButton");

        private void ButtonClicked(object sender) => new StringBuilder()
            .Append($"PGSolutions Better Ribbon\n\n")
            .Append($"Better Ribbon V {Globals.ThisAddIn.VersionNo3}\n")
            .Append($"Ribbon Utilities V {UtilitiesVersion.Format2()}\n")
            .Append($"Ribbon Dispatcher V {DispatcherVersion.Format2()}\n\n")
            .Append($"{BrandingButton.SuperTip}")
            .ToString().MsgBoxShow();

        static Version DispatcherVersion => new RibbonFactory().GetType().Assembly.GetName().Version;
        static Version UtilitiesVersion  => new VbaExportEventArgs(null).GetType().Assembly.GetName().Version;
    }
}
