////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary>The total interface (required to be) exposed externally by RibbonDropDown objects; 
    /// composition of IRibbonCommon, IDropDownItem &amp; IImageableItem</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IRibbonDropDown)]
    public interface IRibbonDropDown : IRibbonCommon {
        /// <summary>Returns the unique (within this ribbon) identifier for this control.</summary>
        [DispId(DispIds.Id)]
        [Description("Returns the unique (within this ribbon) identifier for this control.")]
        new string      Id         { get; }
        /// <summary>Returns the Description string for this control. Only applicable for Menu Items.</summary>
        [Description("Returns the Description string for this control. Only applicable for Menu Items.")]
        [DispId(DispIds.Description)]
        new string Description     { get; }
        /// <summary>Returns the KeyTip string for this control.</summary>
        [Description("Returns the KeyTip string for this control.")]
        [DispId(DispIds.KeyTip)]
        new string KeyTip          { get; }
        /// <summary>Returns the Label string for this control.</summary>
        [Description("Returns the Label string for this control.")]
        [DispId(DispIds.Label)]
        new string Label           { get; }
        /// <summary>Returns the screenTip string for this control.</summary>
        [Description("Returns the screenTip string for this control.")]
        [DispId(DispIds.ScreenTip)]
        new string ScreenTip       { get; }
        /// <summary>Returns the SuperTip string for this control.</summary>
        [Description("Returns the SuperTip string for this control.")]
        [DispId(DispIds.SuperTip)]
        new string SuperTip        { get; }

        /// <summary>TODO</summary>
        [Description("")]
        [DispId(DispIds.IsEnabled)]
        new bool IsEnabled       { get; }
        /// <summary>TODO</summary>
        [Description("")]
    //    [DispId(DispIds.IsVisible)]
        new bool IsVisible       { get; }

        /// <inheritdoc/>
        new void Invalidate();

        /// <summary>Returns the ID of the current selected item.</summary>
        [Description("Returns the ID of the current selected item.")]
        [DispId(DispIds.SelectedItemId)]
        string      SelectedItemId      { get; }
        /// <summary>Returns the index of the current selected item.</summary>
        [Description("Returns the index of the current selected item.")]
        [DispId(DispIds.SelectedItemIndex)]
        int         SelectedItemIndex   { get; }
        /// <summary>Call back for OnAction events from the drop-down ribbon elements.</summary>
        [Description("Call back for OnAction events from the drop-down ribbon elements.")]
        [DispId(DispIds.OnActionDropDown)]
        void        OnActionDropDown(string SelectedId, int SelectedIndex);

        /// <summary>Returns this RibbonDropDown with a new {SelectableItem} in its list.</summary>
        [Description("Returns this RibbonDropDown with a new {SelectableItem} in its list.")]
        [DispId(DispIds.AddItem)]
        IRibbonDropDown AddItem(ISelectableItem SelectableItem);

        /// <summary>Call back for ItemCount events from the drop-down ribbon elements.</summary>
        [Description("Call back for ItemCount events from the drop-down ribbon elements.")]
        [DispId(DispIds.ItemCount)]
        int         ItemCount           { get; }
        /// <summary>Call back for GetItemID events from the drop-down ribbon elements.</summary>
        [Description("Call back for GetItemID events from the drop-down ribbon elements.")]
        [DispId(DispIds.ItemId)]
        string      ItemId(int Index);
        /// <summary>Call back for GetItemLabel events from the drop-down ribbon elements.</summary>
        [Description("Call back for GetItemLabel events from the drop-down ribbon elements")]
        [DispId(DispIds.ItemLabel)]
        string      ItemLabel(int Index);
        /// <summary>Call back for GetItemScreenTip events from the drop-down ribbon elements.</summary>
        [Description("Call back for GetItemScreenTip events from the drop-down ribbon elements.")]
        [DispId(DispIds.ItemScreenTip)]
        string      ItemScreenTip(int Index);
        /// <summary>Call back for GetItemSuperTip events from the drop-down ribbon elements.</summary>
        [Description("Call back for GetItemSuperTip events from the drop-down ribbon elements.")]
        [DispId(DispIds.ItemSuperTip)]
        string      ItemSuperTip(int Index);

        /// <summary>Call back for GetItemSuperTip events from the drop-down ribbon elements.</summary>
        [Description("Call back for GetItemSuperTip events from the drop-down ribbon elements.")]
        [DispId(DispIds.ItemShowImage)]
        bool        ItemShowImage(int Index);
        /// <summary>Call back for GetItemSuperTip events from the drop-down ribbon elements.</summary>
        [Description("Call back for GetItemSuperTip events from the drop-down ribbon elements.")]
        [DispId(DispIds.ItemShowLabel)]
        bool        ItemShowLabel(int Index);
    }
}
