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
    [Guid(Guids.IGalleryModel)]
    public interface IGalleryModel {
        /// <summary>Gets the <see cref="IRibbonControlStrings"/> for this control.</summary>
        IControlStrings Strings {
            [Description("Gets the IControlStrings for this control.")]
            get;
        }

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

        /// <summary>Adds the specified <see cref="ISelectableItem"/> to the available options in the drop-down list.</summary>
        IGalleryModel AddSelectableModel(ISelectableItemModel selectableModel);

        /// <summary>Gets or sets the height in pixels for items.</summary>
        int ItemHeight {
            [Description("Gets or sets the height in pixels for items.")]
            get; set; }

        /// <summary>Gets or sets the width in pixels for items.</summary>
        int ItemWidth{
            [Description("Gets or sets the width in pixels for items.")]
            get; set; }

        /// <summary>Attaches this control-model to the specified ribbon-control as data source and event sink.</summary>
        [Description("Attaches this control-model to the specified ribbon-control as data source and event sink.")]
        IGalleryModel Attach(string controlId);

        /// <summary>.</summary>
        [Description(".")]
        void Detach();

        /// <summary>Queues a request for this control to be refreshed.</summary>
        [Description("Queues a request for this control to be refreshed.")]
        void Invalidate();
    }
}
