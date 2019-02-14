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
    public sealed class RibbonGroupModel : RibbonControlModel<IRibbonGroup>, IRibbonGroupModel {
        public RibbonGroupModel(Func<string, IRibbonGroup> funcViewModel,
                IRibbonControlStrings strings, bool isEnabled, bool isVisible)
        : base(strings, isEnabled, isVisible)
        => FuncViewModel   = funcViewModel;

        #region IActivatable implementation
        private Func<string, IRibbonGroup> FuncViewModel { get; }

        /// <inheritdoc/>
        public IRibbonGroupModel Attach(string controlId) {
            ViewModel = FuncViewModel(controlId);
            ViewModel.Attach(()=>ShowInactive).SetLanguageStrings(Strings);
            Invalidate();
            return this;
        }

        /// <inheritdoc/>
        public override void Invalidate() {
            if (ViewModel != null) {
                base.Invalidate();
            }
        }
        #endregion

        /// <inheritdoc/>
        public bool ShowInactive { get; private set; }

        public void SetShowInactive(bool showInactive) {
            ShowInactive = showInactive;
            ViewModel?.SetShowInactive(ShowInactive);
            ViewModel?.Invalidate();
        }
    }

    public abstract class RibbonControlModel<T> : IRibbonControlModel where T:IRibbonCommon {
        protected RibbonControlModel(IRibbonControlStrings strings, bool isEnabled, bool isVisible) {
            Strings   = strings;
            IsEnabled = isEnabled;
            IsVisible = IsVisible;
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
