////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.RibbonDispatcher.Models {
    using IStrings = IControlStrings;

    public abstract class AbstractRibbonGroupModel : IControlSource, ICanInvalidate {
        protected AbstractRibbonGroupModel(IRibbonViewModel viewModel, string viewModelName, IStrings strings)
        : this(viewModel?.GetControl<GroupVM>(viewModelName), strings) {
        }
        protected AbstractRibbonGroupModel(IGroupVM viewModel, IStrings strings) {
            ViewModel = (viewModel as IActivatable<IControlSource,IGroupVM>)
                      ?.Attach(this);
            Label     = strings.Label;
            ScreenTip = strings.ScreenTip;
            SuperTip  = strings.SuperTip;
            KeyTip    = strings.KeyTip;
        }

        /// <inheritdoc/>
        public string   Label        { get; set; }
        /// <inheritdoc/>
        public string   ScreenTip    { get; set; }
        /// <inheritdoc/>
        public string   SuperTip     { get; set; }
        /// <inheritdoc/>
        public string   KeyTip       { get; set; }
        public bool     IsEnabled    { get; set; } = true;
        public bool     IsVisible    { get; set; } = true;
        public bool     ShowInactive { get; private set; } = true;

        internal IGroupVM ViewModel { get; }

        public void Invalidate() => Invalidate(null);

        internal virtual void Invalidate(Action<IControlVM> action) => ViewModel?.Invalidate(action);

        /// <summary>Set ShowInactive for al- child controls of this ViewModel - even the unattached.</summary>
        /// <param name="showInactive">The <see cref="bool"/> value to be set</param>
        public void SetShowInactive(bool showInactive) {
            ShowInactive = showInactive;
            ViewModel?.Invalidate(c => c.SetShowInactive(ShowInactive));
        }

        public void DetachControls() => ViewModel?.Detach();
    }
}
