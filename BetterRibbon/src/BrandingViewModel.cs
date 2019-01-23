////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using stdole;

using PGSolutions.RibbonDispatcher.ComClasses;
using static Microsoft.Office.Core.RibbonControlSize;

namespace PGSolutions.BetterRibbon {
    internal class BrandingViewModel : AbstractRibbonGroupViewModel {
        [SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "System.Windows.Forms.MessageBox.Show(System.String,System.String,System.Windows.Forms.MessageBoxButtons,System.Windows.Forms.MessageBoxIcon)")]
        [SuppressMessage("Microsoft.Globalization", "CA1300:SpecifyMessageBoxOptions")]
        public BrandingViewModel(IRibbonFactory factory, Func<IPictureDisp> logo) : base(factory) {
            BrandingGroup  = Factory.NewRibbonGroup("BrandingGroup", true);
            BrandingButton = Factory.NewRibbonButton("BrandingButton", true, true, RibbonControlSizeLarge, logo(), false, false);

            BrandingButton.Clicked += OnButtonClicked;
            BrandingButton.Attach();
        }

        public event ClickedEventHandler  ButtonClicked;

        [SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        public RibbonGroup  BrandingGroup  { get; }
        public RibbonButton BrandingButton { get; }

        public void Invalidate() => BrandingButton.Invalidate();

        private void OnButtonClicked(object sender) => ButtonClicked?.Invoke(sender);
    }
}
