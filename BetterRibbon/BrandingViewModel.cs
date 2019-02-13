////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using stdole;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ComClasses;

namespace PGSolutions.BetterRibbon {
    internal class BrandingViewModel : RibbonGroupViewModel {
        public BrandingViewModel(IRibbonFactory factory, Func<IPictureDisp> logo, bool isVisible = true, bool isEnabled = true)
        : base(factory, "BrandingGroup", isVisible, isEnabled) {
            BrandingButton = Factory.NewRibbonButton("BrandingButton", image:logo(), showImage:false, showLabel:false);

            BrandingButton.Clicked += OnButtonClicked;
            BrandingButton.Attach();
        }

        public event ClickedEventHandler  ButtonClicked;

        public RibbonButton BrandingButton { get; }

        public override void Invalidate() {
            BrandingButton.Invalidate();
            base.Invalidate();
        }

        private void OnButtonClicked(object sender) => ButtonClicked?.Invoke(sender);
    }
}
