////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    public abstract class RibbonControlModel<T> : IRibbonControlSource where T:IRibbonCommon {
        protected RibbonControlModel(IRibbonControlStrings strings, bool isEnabled, bool isVisible) {
            Strings    = strings;
            _isEnabled = isEnabled;
            _isVisible = isVisible;
        }

        /// <inheritdoc/>
        public IRibbonControlStrings Strings   { get; protected set; }

        /// <inheritdoc/>
        public T ViewModel { get; set; }

        /// <inheritdoc/>
        public bool IsEnabled {
            get => _isEnabled;
            set { _isEnabled = value; if(ViewModel!=null) ViewModel.IsEnabled = _isEnabled; }
        } private bool _isEnabled = true;

        /// <inheritdoc/>
        public bool IsVisible {
            get => _isVisible;
            set { _isVisible = value; if (ViewModel!=null) ViewModel.IsVisible = _isVisible; }
        } private bool _isVisible = true;

        /// <inheritdoc/>
        public virtual void Invalidate() {
            if (ViewModel != null) { ViewModel.Invalidate(); }
        }
    }
}
