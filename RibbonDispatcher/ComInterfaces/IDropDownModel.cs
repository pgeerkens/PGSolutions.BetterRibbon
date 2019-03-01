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
        event SelectionMadeEventHandler SelectionMade;

        /// <summary>Gets or sets the (zero-based) integer of the selected item. Default value.</summary>
        [DispId(0)]
        int SelectedIndex {
            [Description("Gets or sets the (zero-based) integer of the selected item. Default value.")]
            get; set; }

        #region IActivable implementation
        /// <summary>Attaches this control-model to the specified ribbon-control as data source and event sink.</summary>
        [DispId(1),Description("Attaches this control-model to the specified ribbon-control as data source and event sink.")]
        IDropDownModel Attach(string controlId);

        /// <summary>.</summary>
        [DispId(2),Description(".")]
        void Detach();

        /// <summary>Queues a request for this control to be refreshed.</summary>
        [DispId(3),Description("Queues a request for this control to be refreshed.")]
        void Invalidate();
        #endregion

        #region IControl implementation
        /// <summary>Gets the {IControlStrings} for this control.</summary>
        [DispId(4)]
        IControlStrings Strings {
            [Description("Gets the {IControlStrings} for this control.")]
            get;
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

        /// <summary>Gets or sets the ID of the selected item.</summary>
        [DispId(7)]
        string SelectedId {
            [Description("Gets or sets the ID of the selected item.")]
            get; set; }

        /// <summary>Adds the specified <see cref="ISelectableItem"/> to the available options in the drop-down list.</summary>
        [DispId( 8),Description("Adds the specified ISelectableItem to the available options in the drop-down list.")]
        IDropDownModel ClearList();

        /// <summary>Adds the specified <see cref="ISelectableItem"/> to the available options in the drop-down list.</summary>
        [DispId( 9),Description("Adds the specified ISelectableItem to the available options in the drop-down list.")]
        IDropDownModel AddSelectableModel(IStaticItemVM selectableModel);
    }
}
