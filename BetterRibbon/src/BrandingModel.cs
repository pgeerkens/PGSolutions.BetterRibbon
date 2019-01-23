////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Windows.Forms;

namespace PGSolutions.BetterRibbon {
    internal sealed class BrandingModel {
        public BrandingModel(BrandingViewModel viewModel) =>
            ViewModel = viewModel;

        public void Attach()     => ViewModel.ButtonClicked += ButtonClicked;
        public void Detach()     => ViewModel.ButtonClicked -= ButtonClicked;
        public void Invalidate() => ViewModel.Invalidate();

        private BrandingViewModel ViewModel { get; set; }

        private string VersionNo => GetType().Assembly.GetName().Version.ToString();
        private void ButtonClicked(object sender) =>
            MessageBox.Show("Quack, eh!\n\n" + VersionNo,
                    "PGSolutions - VBA Tools",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
    }
}
