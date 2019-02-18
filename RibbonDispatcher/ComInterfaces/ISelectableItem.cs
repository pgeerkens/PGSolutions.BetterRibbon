////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.InteropServices;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary>TODO</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.ISelectableItem)]
    public interface ISelectableItem {
        /// <summary>TODO</summary>
        string   Id         { get; }
        /// <summary>TODO</summary>
        string   Label      { get; }
        /// <summary>TODO</summary>
        string   ScreenTip  { get; }
        /// <summary>TODO</summary>
        string   SuperTip   { get; }

        /// <summary>Call back for GetItemScreenTip events from the drop-down ribbon elements.</summary>
        bool     ShowImage  { get; }
        /// <summary>Call back for GetItemSuperTip events from the drop-down ribbon elements.</summary>
        bool     ShowLabel  { get; }
    }
}
