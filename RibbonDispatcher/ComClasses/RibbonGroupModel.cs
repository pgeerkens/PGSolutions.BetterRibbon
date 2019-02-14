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
}
