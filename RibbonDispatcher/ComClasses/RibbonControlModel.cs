﻿////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    public abstract class RibbonControlModel<TSource,TCtrl>: IRibbonCommonSource
            where TSource: IRibbonCommonSource
            where TCtrl: class, IRibbonControlVM {
        protected RibbonControlModel(Func<string, IActivatable<TSource, TCtrl>> funcViewModel,
                IRibbonControlStrings strings, bool isEnabled, bool isVisible) {
            AttachToViewModel = (controlId, source) => funcViewModel(controlId)?.Attach(source);
            Strings   = strings;
            IsEnabled = isEnabled;
            IsVisible = isVisible;
        }

        protected Func<string, TSource, TCtrl> AttachToViewModel { get; }

        /// <inheritdoc/>
        public IRibbonControlStrings Strings { get; }

        /// <inheritdoc/>
        public TCtrl ViewModel    { get; set; }

        /// <inheritdoc/>
        public bool  IsEnabled    { get; set; } = true;

        /// <inheritdoc/>
        public bool  IsVisible    { get; set; } = true;

        public bool  ShowInactive { get; private set; } = true;

        /// <inheritdoc/>
        public virtual void Invalidate() => ViewModel?.Invalidate();

        /// <inheritdoc/>
        public virtual void SetShowInactive(bool showInactive) {
            ShowInactive = showInactive;
            ViewModel?.Invalidate();
        }
    }
}
