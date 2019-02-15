////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
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
    public class RibbonGroupModel : RibbonControlModel<RibbonGroupViewModel>, IRibbonGroupModel,
                IRibbonCommonSource {
        public RibbonGroupModel(Func<string, RibbonGroupViewModel> funcViewModel,
                IRibbonControlStrings strings, bool isEnabled, bool isVisible)
        : base(funcViewModel, strings, isEnabled, isVisible)
        { }

        /// <inheritdoc/>
        public IRibbonGroupModel Attach(string controlId) {
            ViewModel = (FuncViewModel(controlId) as IActivatable<RibbonGroupViewModel, IRibbonCommonSource>)
                      ?.Attach(this);
            if (ViewModel != null) {
                ViewModel.Invalidate();
            }
            return this;
        }
    }
}
