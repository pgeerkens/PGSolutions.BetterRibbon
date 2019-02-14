////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    public abstract class RibbonControlModel<T> : IRibbonControlModel where T:IRibbonCommon {
        protected RibbonControlModel(IRibbonControlStrings strings, bool isEnabled, bool isVisible) {
            Strings   = strings;
            IsEnabled = isEnabled;
            IsVisible = isVisible;
        }

        /// <inheritdoc/>
        public IRibbonControlStrings Strings   { get; protected set; }

        /// <inheritdoc/>
        public bool                  IsEnabled { get; set; } = true;

        /// <inheritdoc/>
        public bool                  IsVisible { get; set; } = true;

        /// <inheritdoc/>
        public T    ViewModel { get; set; }

        /// <inheritdoc/>
        public virtual void Invalidate() {
            if (ViewModel != null) {
                ViewModel.IsEnabled = IsEnabled;
                ViewModel.IsVisible = IsVisible;

                ViewModel.Invalidate();
            }
        }
    }
}
