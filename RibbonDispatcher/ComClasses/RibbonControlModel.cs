////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////

using System;
using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    public abstract class RibbonControlModel<T> : IRibbonCommonSource where T:IRibbonCommon {
        protected RibbonControlModel(Func<string,T> funcViewModel,
                IRibbonControlStrings strings, bool isEnabled, bool isVisible) {
            FuncViewModel = funcViewModel;
            Strings       = strings;
            IsEnabled     = isEnabled;
            IsVisible     = isVisible;

            //Invalidate();
        }

        protected Func<string, T> FuncViewModel { get; }

        /// <inheritdoc/>
        public IRibbonControlStrings Strings { get; protected set; }

        /// <inheritdoc/>
        public T ViewModel { get; set; }

        /// <inheritdoc/>
        public bool IsEnabled { get; set; } = true;

        /// <inheritdoc/>
        public bool IsVisible { get; set; } = true;

        public bool ShowInactive { get; private set; } = true;

        /// <inheritdoc/>
        public virtual void Invalidate() {
            if (ViewModel != null) { ViewModel.Invalidate(); }
        }

        public virtual void SetShowInactive(bool showInactive) {
            ShowInactive = showInactive;
            ViewModel?.Invalidate();
        }
    }
}
