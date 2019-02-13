////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    /// <summary></summary>
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable")]
    [Description("")]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IRibbonGroupModel))]
    [Guid(Guids.RibbonGroupModel)]
    public sealed class RibbonGroupModel : IRibbonGroupModel, IBooleanSource {
        public RibbonGroupModel(Func<string, ICustomRibbonGroup> factory,
                IRibbonControlStrings strings, bool isEnabled, bool isVisible) {
            Factory = factory;
            Strings = strings;
        }

        public IRibbonControlStrings Strings { get; }
        public bool IsEnabled { get; set; } = true;
        public bool IsVisible { get; set; } = true;

        #region IActivatable implementation
        /// <inheritdoc/>
        public ICustomRibbonGroup ViewModel { get; set; }

        private Func<string,ICustomRibbonGroup> Factory { get; }

        /// <inheritdoc/>
        public IRibbonGroupModel Attach(string controlId) {
            ViewModel = Factory(controlId);
            ViewModel.Attach();
            Invalidate();
            return this;
        }

        /// <inheritdoc/>
        public void Invalidate() {
            if (ViewModel != null) {
                ViewModel.IsEnabled = IsEnabled;
                ViewModel.IsVisible = IsVisible;

                ViewModel.Invalidate();
            }
        }
        #endregion

        #region IToggleable implementation
        public event ToggledEventHandler Toggled;

        private void OnToggled(object sender, bool showInactive)
        => Toggled?.Invoke(sender, ShowInactive = showInactive);

        /// <inheritdoc/>
        public bool ShowInactive {
            get => _showInactive;
            set { _showInactive = value; ViewModel?.SetShowInactive(value); ViewModel?.Invalidate(); }
        }
        bool _showInactive = false;

        public bool Getter() => ShowInactive;
        #endregion
    }
}
