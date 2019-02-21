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
    public class RibbonGroupModel : RibbonControlModel<IRibbonCommonSource,RibbonGroupViewModel>,
            IRibbonGroupModel, IRibbonCommonSource {
        public RibbonGroupModel(Func<string,RibbonGroupViewModel> funcViewModel,
                IRibbonControlStrings strings, bool isEnabled, bool isVisible,
                IRibbonCommonSource groupMaster)
        : base(funcViewModel, strings, isEnabled, isVisible)
        => GroupMaster = groupMaster;

        /// <inheritdoc/>
        public IRibbonGroupModel Attach(string controlId) {
            ViewModel = AttachToViewModel(controlId, this);
            if (ViewModel != null) {
                ViewModel.Invalidate();
            }
            return this;
        }

        public override void SetShowInactive(bool showInactive)
        => GroupMaster.SetShowInactive(showInactive);

        public void Detach() => ViewModel.Detach();

        private IRibbonCommonSource GroupMaster { get; }
    }
}
