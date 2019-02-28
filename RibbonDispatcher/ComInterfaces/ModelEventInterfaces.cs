////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

using Microsoft.Office.Core;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    public enum ERibbonControlSize {
        SizeSmall = RibbonControlSize.RibbonControlSizeRegular,
        SizeLarge = RibbonControlSize.RibbonControlSizeLarge
    }

    /// <summary>TODO</summary>
    /// <remarks>Must be an IDispatch interface because an event source.</remarks>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    [Guid(Guids.IClickedEvent)]
    public interface IClickedEvent {
        /// <summary>Fired when the associated control is clicked by the user.</summary>
        [Description("Fired when the associated control is clicked by the user.")]
        void Clicked(IRibbonControl control);
    }

    /// <summary>TODO</summary>
    /// <remarks>Must be an IDispatch interface because an event source.</remarks>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    [Guid(Guids.IToggledEvent)]
    public interface IToggledEvent {
        /// <summary>Fired when the associated control is toggled by the user.</summary>
        [Description("Fired when the associated control is toggled by the user.")]
        void Toggled(IRibbonControl control, bool IsPressed);
    }

    /// <summary>TODO</summary>
    /// <remarks>Must be an IDispatch interface because an event source.</remarks>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    [Guid(Guids.IEditedEvent)]
    public interface IEditedEvent {
        /// <summary>Fired when the associated control is clicked by the user.</summary>
        [Description("Fired when the associated control is edited by the user.")]
        void Edited(IRibbonControl control, string text);
    }

    /// <summary>TODO</summary>
    /// <remarks>Must be an IDispatch interface because an event source.</remarks>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    [Guid(Guids.ISelectionMadeEvent)]
    public interface ISelectionMadeEvent {
        /// <summary>Fired when the associated control has an item selection made by the user.</summary>
        [Description("Fired when the associated control has an item selection made by the user.")]
        void SelectionMade(IRibbonControl control, string selectedId, int selectedIndex);
    }
}
