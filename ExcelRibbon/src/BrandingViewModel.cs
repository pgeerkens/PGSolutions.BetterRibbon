////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using System.Windows.Forms;
using stdole;

using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonDispatcher.ComInterfaces;
using static PGSolutions.RibbonDispatcher.ComInterfaces.RdControlSize;

namespace PGSolutions.ExcelRibbon {
    internal class BrandingViewModel : AbstractRibbonGroupViewModel {
        [SuppressMessage("Microsoft.Globalization", "CA1300:SpecifyMessageBoxOptions")]
        [SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "System.Windows.Forms.MessageBox.Show(System.String,System.String,System.Windows.Forms.MessageBoxButtons)")]
        public BrandingViewModel(IRibbonFactory factory, Func<IPictureDisp> logo) : base(factory) {
            BrandingGroup  = Factory.NewRibbonGroup("BrandingGroup", true);
            BrandingButton = Factory.NewRibbonButton("BrandingButton", true, true, rdLarge, logo(), false, false);

            BrandingButton.Clicked += () =>
                MessageBox.Show("Quack, eh!\n\n" + typeof(BrandingViewModel).Assembly.GetName().Version.ToString(),
                        "PGSolutions - VBA Tools", MessageBoxButtons.OK);
            BrandingButton.Attach(null);
        }

        public RibbonGroup  BrandingGroup  { get; }
        public RibbonButton BrandingButton { get; }
    }
}
