////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.RibbonDispatcher.Models {
    /// <summary>The COM visible Model for Ribbon Group controls.</summary>
    [SuppressMessage("Microsoft.Interoperability","CA1409:ComVisibleTypesShouldBeCreatable")]
    [Description("The COM visible Model for Ribbon Group controls.")]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IGroupModel))]
    [Guid(Guids.GroupModel)]
    public class GroupModel : ControlModel<IControlSource,IGroupVM>,
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
