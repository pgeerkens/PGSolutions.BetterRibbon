////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Linq;
using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    using IStrings = IRibbonControlStrings;

    [CLSCompliant(false)]
    public abstract class AbstractRibbonTabModel {
        protected AbstractRibbonTabModel(AbstractRibbonViewModel viewModel) {
            ViewModel = viewModel;
        }

        public void Initialize(IReadOnlyList<IInvalidate> models) => Models = models;

        public    AbstractRibbonViewModel    ViewModel          { get; }

        protected IReadOnlyList<IInvalidate> Models             { get; private set; }

        protected List<CustomButtonsGroupModel> CustomButtonsModel
        => Models.OfType<CustomButtonsGroupModel>().ToList();

        public void Invalidate() { foreach (var model in Models) { model?.Invalidate(); } }

        /// <inheritdoc/>
        internal void DetachProxy(string controlId) => GetControl<IRibbonCommon>(controlId)?.Detach();

        private TControl GetControl<TControl>(string controlId) where TControl : class, IRibbonCommon
        => ViewModel.RibbonFactory.GetControl<TControl>(controlId);

        public void DetachCustomControls()
        => CustomButtonsModel.ForEach(model => model.DetachControls());

        protected IStrings GetStrings(string id)
        => ViewModel.RibbonFactory.ResourceManager.GetControlStrings(id);

        public RibbonButtonModel NewRibbonButtonModel(string id, EventHandler handler,
                bool isEnabled, bool isVisible, ImageObject image) {
            var model = NewRibbonButtonModel(GetStrings(id), image, isEnabled, isVisible);

            model?.Attach(id);
            model.Clicked += handler;
            return model;
        }

        public RibbonToggleModel NewRibbonToggleModel(string id, ToggledEventHandler handler, bool isEnabled,
                bool isVisible, ImageObject image) {
            var model = NewRibbonToggleModel(GetStrings(id), image, isEnabled, isVisible);

            model?.Attach(id);
            model.Toggled += handler;
            return model;
        }

        public RibbonDropDownModel NewRibbonDropDownModel(string id, SelectedEventHandler handler, bool isEnabled,
                bool isVisible) {
            var model = NewRibbonDropDownModel(GetStrings(id), isEnabled, isVisible);

            model?.Attach(id);
            model.SelectionMade += handler;
            return model;
        }

        /// <inheritdoc/>
        public RibbonButtonModel NewRibbonButtonModel(IStrings strings,
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
        => new RibbonGroupModel(ViewModel.RibbonFactory.GetControl<RibbonGroupViewModel>, strings, isEnabled, isVisible);
    }
}
