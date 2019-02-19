////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    using IStrings = IRibbonControlStrings;

    [CLSCompliant(false)]
    public abstract class AbstractRibbonGroupModel : IRibbonCommonSource, IInvalidate {
        protected AbstractRibbonGroupModel(RibbonGroupViewModel viewModel) {
            ViewModel = (viewModel as IActivatable<IRibbonCommonSource,RibbonGroupViewModel>)
                      ?.Attach(this);
            Strings   = GetStrings(ViewModel.Id);
        }

        public bool     IsEnabled    { get; set; } = true;
        public bool     IsVisible    { get; set; } = true;
        public bool     ShowInactive { get; private set; } = true;
        public IStrings Strings      { get; private set; }

        protected RibbonGroupViewModel ViewModel { get; }

        public void Invalidate() => Invalidate(null);

        public virtual void Invalidate(Action<IActivatable> action) => ViewModel.Invalidate(action);

        /// <summary>Set ShowInactive for al- child controls of this ViewModel - even the unattached.</summary>
        /// <param name="showInactive">The <see cref="bool"/> value to be set</param>
        public void SetShowInactive(bool showInactive) {
            ShowInactive = showInactive;
            ViewModel.Invalidate(c => c.SetShowInactive(ShowInactive));
        }

        public void DetachControls() => ViewModel?.DetachControls();

        public TControl GetControl<TControl>(string controlId) where TControl : class, IRibbonCommon
        => ViewModel.GetControl<TControl>(controlId);

        protected RibbonButtonModel NewButtonModel(string id, EventHandler handler,
                bool isEnabled, bool isVisible, ImageObject image) {
            var model = new RibbonButtonModel(GetControl<RibbonButton>, GetStrings(id), image, isEnabled, isVisible);

            ViewModel.Add<IRibbonButtonSource>(ViewModel.Factory.NewRibbonButton(id));
            model?.Attach(id);
            model.Clicked += handler;
            return model;
        }

        internal RibbonButtonModel NewButtonModel(IStrings strings, ImageObject image,
                bool isEnabled, bool isVisible) {
            var model = new RibbonButtonModel(GetControl<RibbonButton>, strings, image, isEnabled, isVisible);

            model.SetShowInactive(false);
            model.Invalidate();
            return model;
        }

        protected RibbonToggleModel NewToggleModel(string id, ToggledEventHandler handler, bool isEnabled,
                bool isVisible, ImageObject image) {
            var model = new RibbonToggleModel(GetControl<RibbonCheckBox>, GetStrings(id), image, isEnabled, isVisible);

            ViewModel.Add<IRibbonToggleSource>(ViewModel.Factory.NewRibbonToggle(id));
            model?.Attach(id);
            model.Toggled += handler;
            return model;
        }

        public RibbonToggleModel NewToggleModel(IStrings strings, ImageObject image,
                bool isEnabled, bool isVisible) {
            var model = new RibbonToggleModel(GetControl<RibbonCheckBox>, strings, image, isEnabled, isVisible);

            model.SetShowInactive(false);
            model.Invalidate();
            return model;
        }

        protected RibbonDropDownModel NewDropDownModel(string id, SelectedEventHandler handler, bool isEnabled,
                bool isVisible) {
            var model = new RibbonDropDownModel(GetControl<RibbonDropDown>, GetStrings(id), isEnabled, isVisible);

            ViewModel.Add<IRibbonDropDownSource>(ViewModel.Factory.NewRibbonDropDown(id));
            model?.Attach(id);
            model.SelectionMade += handler;
            return model;
        }

        public RibbonDropDownModel NewDropDownModel(IStrings strings,
                bool isEnabled, bool isVisible) {
            var model = new RibbonDropDownModel(GetControl<RibbonDropDown>, strings, isEnabled, isVisible);

            model.SetShowInactive(false);
            model.Invalidate();
            return model;
        }

        protected IStrings GetStrings(string id)
        => ViewModel.Factory.ResourceManager.GetControlStrings(id);
    }
}
