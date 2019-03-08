////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary></summary>
    [Description("")]
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IDynamicMenuModel)]
    public interface IDynamicMenuModel {
        #region IActivable implementation
        /// <summary>Attaches this control-model to the specified ribbon-control as data source and event sink.</summary>
        [DispId(1),Description("Attaches this control-model to the specified ribbon-control as data source and event sink.")]
        IDynamicMenuModel Attach(string controlId);

        /// <summary>.</summary>
        [DispId(2),Description(".")]
        void Detach();

        /// <summary>Queues a request for this control to be refreshed.</summary>
        [DispId(3),Description("Queues a request for this control to be refreshed.")]
        void Invalidate();
        #endregion

        #region IControl implementation
        /// <summary>Gets the <see cref="IControlStrings"/> for this control.</summary>
        [DispId(4)]
        string Label {
            [Description("Gets the IControlStrings for this control.")]
            get; set;
        }
        /// <summary>Gets the ScreenTip (concise hover-help) for this control.</summary>
        [DispId(17)]
        string ScreenTip {
            [Description("Gets the ScreenTip (concise hover-help) for this control.")]
            get; set;
        }
        /// <summary>Gets the SuperTip (expanded hover-help) for this control.</summary>
        [DispId(18)]
        string SuperTip {
            [Description("Gets the SuperTip (expanded hover-help) for this control.")]
            get; set;
        }
        /// <summary>Gets the KeyTip (keyboard shortcut) for this control.</summary>
        [DispId(19)]
        string KeyTip {
            [Description("Gets the KeyTip (keyboard shortcut) for this control.")]
            get; set;
        }
        /// <summary>Gets or sets whether the control is enabled.</summary>
        [DispId(5)]
        bool IsEnabled {
            [Description("Gets or sets whether the control is enabled.")]
            get; set;
        }
        /// <summary>Gets or sets whether the control is visible.</summary>
        [DispId(6)]
        bool IsVisible {
            [Description("Gets or sets whether the control is visible.")]
            get; set;
        }
        #endregion
    }
}
