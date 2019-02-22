////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ComClasses.ViewModels;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    /// <summary></summary>
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable")]
    [Description("")]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IRibbonGroupModel))]
    [Guid(Guids.GroupModel)]
    public class GroupModel : RibbonControlModel<IRibbonCommonSource,GroupVM>,
            IRibbonGroupModel, IRibbonCommonSource {
        public GroupModel(Func<string,GroupVM> funcViewModel,
                IRibbonControlStrings strings, bool isEnabled, bool isVisible)
        : base(funcViewModel, strings, isEnabled, isVisible)
        { }

        /// <inheritdoc/>
        public IRibbonGroupModel Attach(string controlId) {
            ViewModel = AttachToViewModel(controlId, this);
            if (ViewModel != null) {
                ViewModel.Invalidate();
            }
            return this;
        }

        public override void SetShowInactive(bool showInactive)
        => ViewModel.SetShowInactive(showInactive);

        public void Detach() => ViewModel.Detach();
    }
}
