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
    [Guid(Guids.IDropDownModel)]
    public interface IDropDownModel {
        /// <summary>Gets the <see cref="IRibbonControlStrings"/> for this control.</summary>
        IControlStrings Strings {
            [DispId( 1),Description("Gets the IControlStrings for this control.")]
            get;
        }

        /// <summary>Gets or sets whether the control is enabled.</summary>
        [DispId(2)]
        bool IsEnabled {
            [Description("Gets or sets whether the control is enabled.")]
            get; set; }

        /// <summary>Gets or sets whether the control is visible.</summary>
        [DispId(3)]
        bool IsVisible {
            [Description("Gets or sets whether the control is visible.")]
            get; set; }

        /// <summary>Gets or sets the (zero-based) integer of the selected item.</summary>
        [DispId(4)]
        int SelectedIndex {
            [Description("Gets or sets the (zero-based) integer of the selected item.")]
            get; set; }

        /// <summary>Adds the specified <see cref="ISelectableItem"/> to the available options in the drop-down list.</summary>
        [DispId( 5),Description("Adds the specified ISelectableItem to the available options in the drop-down list.")]
        IDropDownModel AddSelectableModel(IStaticItemVM selectableModel);

        /// <summary>Attaches this control-model to the specified ribbon-control as data source and event sink.</summary>
        [DispId( 6),Description("Attaches this control-model to the specified ribbon-control as data source and event sink.")]
        IDropDownModel Attach(string controlId);

        /// <summary>.</summary>
        [DispId( 7),Description(".")]
        void Detach();

        /// <summary>Queues a request for this control to be refreshed.</summary>
        [DispId( 8),Description("Queues a request for this control to be refreshed.")]
        void Invalidate();
    }
}
