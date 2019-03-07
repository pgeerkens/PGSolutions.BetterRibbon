////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.RibbonDispatcher.Models {
    using IStrings = IControlStrings;

    public abstract class ControlModel<TSource,TCtrl>: IControlSource
            where TSource: IControlSource
            where TCtrl: class,IControlVM {
        protected ControlModel(Func<string, IActivatable<TSource,TCtrl>> funcViewModel, IStrings strings) {
            AttachToViewModel = (controlId, source) => funcViewModel(controlId)?.Attach(source);
            Label     = strings?.Label;
            ScreenTip = strings?.ScreenTip;
            SuperTip  = strings?.SuperTip;
            KeyTip    = strings?.KeyTip;
        }

        protected Func<string, TSource, TCtrl> AttachToViewModel { get; }

        public virtual void Detach() {
            ViewModel.Detach();
            ViewModel.Invalidate();
            ViewModel = default;
        }

        /// <inheritdoc/>
        public string  Label      { get; set; }
        /// <inheritdoc/>
        public string  ScreenTip  { get; set; }
        /// <inheritdoc/>
        public string  SuperTip   { get; set; }
        /// <inheritdoc/>
        public string  KeyTip     { get; set; }

        /// <inheritdoc/>
        public TCtrl    ViewModel { get; set; }

        /// <inheritdoc/>
        public bool     IsEnabled { get; set; } = true;

        /// <inheritdoc/>
        public bool     IsVisible { get; set; } = true;

        /// <inheritdoc/>
        public virtual void Invalidate() => ViewModel?.Invalidate();

        /// <inheritdoc/>
        public virtual void SetShowInactive(bool showInactive) => ViewModel?.Invalidate();
    }
}
