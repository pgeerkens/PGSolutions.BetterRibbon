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
    [CLSCompliant(false)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IRibbonDropDownModel)]
    public interface IRibbonDropDownModel {
        /// <summary>Gets or sets whether the control is enabled.</summary>
        bool IsEnabled {
            [Description("Gets or sets whether the control is enabled.")]
            get; set; }
        /// <summary>Gets or sets whether the control is visible.</summary>
        bool IsVisible {
            [Description("Gets or sets whether the control is visible.")]
            get; set; }

        /// <summary>Gets or sets the (zero-based) integer of the selected item.</summary>
        int SelectedIndex {
            [Description("Gets or sets the (zero-based) integer of the selected item.")]
            get; set; }

        /// <summary>Adds the specified {ISelectableItem} to the available options in the drop-down list.</summary>
        [Description("Adds the specified ISelectableItem to the available options in the drop-down list.")]
        IRibbonDropDownModel AddItem(ISelectableItem SelectableItem);

        /// <summary>Attaches this control-model to the specified ribbon-control as data source and event sink.</summary>
        [Description("Attaches this control-model to the specified ribbon-control as data source and event sink.")]
        IRibbonDropDownModel Attach(string controlId);

        /// <summary>Queues a request for this control to be refreshed.</summary>
        [Description("Queues a request for this control to be refreshed.")]
        void Invalidate();
    }
}
