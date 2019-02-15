////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Collections;
using System.Collections.Generic;
using PGSolutions.RibbonDispatcher.ComClasses;
using PGSolutions.RibbonDispatcher.ComInterfaces;
using stdole;

namespace PGSolutions.BetterRibbon {
    internal sealed class CustomButtonsModel : IRibbonCommonSource {
        public CustomButtonsModel(RibbonGroupViewModel viewModel) {
            ViewModel = viewModel;
            Models = new ModelsX();

            (ViewModel as IActivatable<IRibbonGroup, IRibbonCommonSource>).Attach(this);

            Invalidate();
        }

        public bool IsEnabled    { get; set; } = true;
        public bool IsVisible    { get; set; } = true;
        public bool ShowInactive { get; set; } = false;
        public IRibbonControlStrings Strings { get; }

        private RibbonGroupViewModel ViewModel { get; set; }

        public void   Invalidate() => ViewModel.Invalidate();

        public void SetShowInactive(bool showInactive) => ViewModel.SetShowInactive(showInactive);

        public TControl GetControl<TControl>(string controlId) where TControl:class,IRibbonCommon
        => ViewModel.GetControl<TControl>(controlId);

        public IRibbonButtonModel NewButtonModel(IRibbonControlStrings strings,
                IPictureDisp image = null, bool isEnabled = true, bool isVisible = true) {
            var model = Models.Add(new RibbonButtonModel(id => GetControl<RibbonButton>(id),
                    strings, image, isEnabled, isVisible));
            model.SetShowInactive(false);
            model.Invalidate();
            return model;
        }

        public IRibbonButtonModel NewButtonModel(IRibbonControlStrings strings,
                string imageMso = null, bool isEnabled = true, bool isVisible = true) {
            var model = Models.Add(new RibbonButtonModel(id => GetControl<RibbonButton>(id),
                    strings, imageMso, isEnabled, isVisible));
            model.SetShowInactive(false);
            model.Invalidate();
            return model;
        }

        public IRibbonToggleModel NewToggleModel(IRibbonControlStrings strings,
                IPictureDisp image = null, bool isEnabled = true, bool isVisible = true) {
            var model = Models.Add(new RibbonToggleModel(id => GetControl<RibbonCheckBox>(id),
                    strings, image, isEnabled, isVisible));
            model.SetShowInactive(false);
            model.Invalidate();
            return model;
        }

        public IRibbonToggleModel NewToggleModel(IRibbonControlStrings strings,
                string imageMso = null, bool isEnabled = true, bool isVisible = true) {
            var model = Models.Add(new RibbonToggleModel(id => GetControl<RibbonCheckBox>(id),
                    strings, imageMso, isEnabled, isVisible));
            model.SetShowInactive(false);
            model.Invalidate();
            return model;
        }

        public IRibbonDropDownModel NewDropDownModel(IRibbonControlStrings strings,
                bool isEnabled = true, bool isVisible = true) {
            var model = Models.Add(new RibbonDropDownModel(id => GetControl<RibbonDropDown>(id),
                    strings, isEnabled, isVisible));
            model.SetShowInactive(false);
            model.Invalidate();
            return model;
        }

        private ModelsX Models { get; }

        internal class ModelsX : IEnumerable {
            public ModelsX() => List = new List<object>();

            private List<object> List { get; }

            public TControl Add<TControl>(TControl control) { List.Add(control); return control; }

            public IEnumerator GetEnumerator() => List.GetEnumerator();
        }
    }
}
