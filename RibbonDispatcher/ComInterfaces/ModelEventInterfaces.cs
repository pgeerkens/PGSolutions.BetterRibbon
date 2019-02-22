////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary>TODO</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    [Guid(Guids.IClickedEvents)]
    public interface IClickedEvents {
        /// <summary>Fired when the associated control is clicked by the user.</summary>
        [Description("Fired when the associated control is clicked by the user.")]
        void Clicked(object sender, EventArgs e);
    }

    /// <summary>TODO</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    [Guid(Guids.IToggledEvents)]
    public interface IToggledEvents {
        /// <summary>Fired when the associated control is toggled by the user.</summary>
        [Description("Fired when the associated control is toggled by the user.")]
        void Toggled(object sender, bool IsPressed);
    }

    /// <summary>TODO</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    [Guid(Guids.ISelectedEvents)]
    public interface ISelectedEvents {
        /// <summary>Fired when the associated control has an item selection made by the user.</summary>
        [Description("Fired when the associated control has an item selection made by the user.")]
        void Selected(object sender, int ItemIndex);
    }

    /// <summary>TODO</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    [Guid(Guids.IEditedEvents)]
    public interface IEditedEvents {
        /// <summary>Fired when the associated control is clicked by the user.</summary>
        [Description("Fired when the associated control is clicked by the user.")]
        void Edited(object sender, string text);
    }
}
