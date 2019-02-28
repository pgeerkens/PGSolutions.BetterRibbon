////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary></summary>
    [Description("")]
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IMenuModel)]
    public interface IMenuModel {
        /// <summary>Gets the <see cref="IControlStrings"/> for this control.</summary>
        [DispId(1)]
        IControlStrings Strings {
            [Description("Gets the IControlStrings for this control.")]
            get;
        }

        /// <summary>Gets or sets whether the control is enabled.</summary>
        [DispId(2)]
        bool IsEnabled {
            [Description("Gets or sets whether the control is enabled.")]
            get; set;
        }
        /// <summary>Gets or sets whether the control is visible.</summary>
        [DispId(3)]
        bool IsVisible {
            [Description("Gets or sets whether the control is visible.")]
            get; set;
        }

        /// <summary>Attaches this control-model to the specified ribbon-control as data source and event sink.</summary>
        [Description("Attaches this control-model to the specified ribbon-control as data source and event sink.")]
        [DispId(4)]
        IMenuModel Attach(string controlId);

        /// <summary>.</summary>
        [Description(".")]
        [DispId(5)]
        void Detach();

        /// <summary>Queues a request for this control to be refreshed.</summary>
        [Description("Queues a request for this control to be refreshed.")]
        [DispId(6)]
        void Invalidate();
    }
}
