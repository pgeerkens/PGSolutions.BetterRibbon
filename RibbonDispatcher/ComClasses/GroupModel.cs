////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ComClasses.ViewModels;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    /// <summary></summary>
    [Description("")]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IGroupModel))]
    [Guid(Guids.GroupModel)]
    public class GroupModel : ControlModel<IControlSource,GroupVM>,
            IGroupModel, IControlSource {
        public GroupModel(Func<string,GroupVM> funcViewModel, IControlStrings strings)
        : base(funcViewModel, strings)
        { }

        /// <inheritdoc/>
        public IGroupModel Attach(string controlId) {
            ViewModel = AttachToViewModel(controlId, this);
            if (ViewModel != null) {
                ViewModel.Invalidate();
            }
            return this;
        }

        /// <inheritdoc/>
        public override void SetShowInactive(bool showInactive)
        => ViewModel.SetShowInactive(showInactive);

        /// <inheritdoc/>
        public bool ShowInactive { get; }
    }
}
