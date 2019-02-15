////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    public abstract class RibbonControlModel<T> : IRibbonCommonSource where T:IRibbonCommon {
        protected RibbonControlModel(IRibbonControlStrings strings, bool isEnabled, bool isVisible) {
            Strings   = strings;

            Invalidate();
        }

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

        public void SetShowInactive(bool showInactive) {
            ShowInactive = showInactive;
            ViewModel?.Invalidate();
        }
    }
}
