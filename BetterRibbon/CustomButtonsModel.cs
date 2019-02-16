////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonDispatcher.ComInterfaces;
using stdole;

namespace PGSolutions.BetterRibbon {
    internal sealed class CustomButtonsModel : IRibbonCommonSource {
        public CustomButtonsModel(RibbonGroupViewModel viewModel) {
            ViewModel = (viewModel as IActivatable<RibbonGroupViewModel,IRibbonCommonSource>)
                      .Attach(this);

            Invalidate();
        }

        public bool IsEnabled    { get; set; } = true;
        public bool IsVisible    { get; set; } = true;
        public bool ShowInactive { get; set; } = false;
        public IRibbonControlStrings Strings { get; }

        internal RibbonGroupViewModel ViewModel { get; set; }

        public void   Invalidate() => ViewModel.Invalidate();

        internal void DetachControls() => ViewModel?.DetachControls();

        public void SetShowInactive(bool showInactive) => ViewModel.SetShowInactive(showInactive);

        public TControl GetControl<TControl>(string controlId) where TControl:class,IRibbonCommon
        => ViewModel.GetControl<TControl>(controlId);

        public IRibbonButtonModel NewButtonModel(IRibbonControlStrings strings,
                IPictureDisp image = null, bool isEnabled = true, bool isVisible = true) {
            var model = new RibbonButtonModel(id => GetControl<RibbonButton>(id),
                    strings, image, isEnabled, isVisible);
            model.SetShowInactive(false);
            model.Invalidate();
            return model;
        }

        public IRibbonButtonModel NewButtonModel(IRibbonControlStrings strings,
                string imageMso = null, bool isEnabled = true, bool isVisible = true) {
            var model = new RibbonButtonModel(id => GetControl<RibbonButton>(id),
                    strings, imageMso, isEnabled, isVisible);
            model.SetShowInactive(false);
            model.Invalidate();
            return model;
        }

        public IRibbonToggleModel NewToggleModel(IRibbonControlStrings strings,
                IPictureDisp image = null, bool isEnabled = true, bool isVisible = true) {
            var model = new RibbonToggleModel(id => GetControl<RibbonCheckBox>(id),
                    strings, image, isEnabled, isVisible);
            model.SetShowInactive(false);
            model.Invalidate();
            return model;
        }

        public IRibbonToggleModel NewToggleModel(IRibbonControlStrings strings,
                string imageMso = null, bool isEnabled = true, bool isVisible = true) {
            var model = new RibbonToggleModel(id => GetControl<RibbonCheckBox>(id),
                    strings, imageMso, isEnabled, isVisible);
            model.SetShowInactive(false);
            model.Invalidate();
            return model;
        }

        public IRibbonDropDownModel NewDropDownModel(IRibbonControlStrings strings,
                bool isEnabled = true, bool isVisible = true) {
            var model = new RibbonDropDownModel(id => GetControl<RibbonDropDown>(id),
                    strings, isEnabled, isVisible);
            model.SetShowInactive(false);
            model.Invalidate();
            return model;
        }
    }
}
