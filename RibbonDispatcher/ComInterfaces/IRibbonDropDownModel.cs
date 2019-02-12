////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary></summary>
    [Description("")]
    [ComVisible(true)]
    [CLSCompliant(false)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IRibbonDropDownModel)]
    public interface IRibbonDropDownModel {
        [SuppressMessage( "Microsoft.Design", "CA1009:DeclareEventHandlersCorrectly", Justification="EventArgs<T> is unknown to COM.")]
        event SelectedEventHandler SelectionMade;

        bool IsEnabled { get; set; }
        bool IsVisible { get; set; }

        int  SelectedIndex { get; set; }

        IRibbonDropDownModel AddItem(ISelectableItem SelectableItem);

        IRibbonDropDownModel Attach(string controlId);

        void Invalidate();
    }
}
