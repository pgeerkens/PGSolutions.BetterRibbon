////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.BetterRibbon {
    internal sealed class CustomButtonsModel : IRibbonCommonSource {
        public CustomButtonsModel(RibbonGroupViewModel viewModel) {
            ViewModel = viewModel;

            (ViewModel as IActivatable<IRibbonGroup, IRibbonCommonSource>).Attach(this);
        }

        public bool IsEnabled    { get; set; }
        public bool IsVisible    { get; set; }
        public bool ShowInactive { get; set; }
        public IRibbonControlStrings Strings { get; }

        private RibbonGroupViewModel ViewModel { get; set; }

        public void   Invalidate() => ViewModel.Invalidate();

        public TControl GetControl<TControl>(string controlId) where TControl:class,IRibbonCommon
        => ViewModel.GetControl<TControl>(controlId);
    }
}
