////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    using IStrings = IRibbonControlStrings;

    [CLSCompliant(false)]
    public abstract class AbstractRibbonTabModel {
        protected AbstractRibbonTabModel(AbstractRibbonViewModel viewModel, IReadOnlyList<IInvalidate> models) {
            ViewModel = viewModel;
            Models    = models;
        }

        public    AbstractRibbonViewModel    ViewModel          { get; }

        protected IReadOnlyList<IInvalidate> Models             { get; }

        protected abstract AbstractRibbonGroupModel CustomButtons1Model { get; }

        public void Invalidate() { foreach (var model in Models) { model?.Invalidate(); } }

        /// <inheritdoc/>
        internal void DetachProxy(string controlId) => GetControl<IRibbonCommon>(controlId).Detach();

        private TControl GetControl<TControl>(string controlId) where TControl : class, IRibbonCommon
        => ViewModel.RibbonFactory.GetControl<TControl>(controlId);

        /// <inheritdoc/>
        public IRibbonButtonModel NewRibbonButtonModel(IStrings strings,
                ImageObject image, bool isEnabled, bool isVisible) {
            var model = new RibbonButtonModel(GetControl<RibbonButton>, strings, image, isEnabled, isVisible);

            model.SetShowInactive(false);
            model.Invalidate();
            return model;
        }

        /// <inheritdoc/>
        public RibbonToggleModel NewRibbonToggleModel(IStrings strings, ImageObject image,
                bool isEnabled, bool isVisible) {
            var model = new RibbonToggleModel(GetControl<RibbonCheckBox>, strings, image, isEnabled, isVisible);

            model.SetShowInactive(false);
            model.Invalidate();
            return model;
        }

        /// <inheritdoc/>
        public RibbonDropDownModel NewRibbonDropDownModel(IStrings strings,
                bool isEnabled, bool isVisible) {
            var model = new RibbonDropDownModel(GetControl<RibbonDropDown>, strings, isEnabled, isVisible);

            model.SetShowInactive(false);
            model.Invalidate();
            return model;
        }

        /// <inheritdoc/>
        public RibbonGroupModel NewRibbonGroupModel(IStrings strings, bool isEnabled, bool isVisible)
        => new RibbonGroupModel(ViewModel.RibbonFactory.GetControl<RibbonGroupViewModel>, strings, isEnabled, isVisible, CustomButtons1Model);
    }
}
