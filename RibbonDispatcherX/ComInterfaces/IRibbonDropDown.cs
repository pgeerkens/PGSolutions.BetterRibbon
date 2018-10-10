////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace PGSolutions.RibbonDispatcher.AbstractCOM {
    /// <summary>The total interface (required to be) exposed externally by RibbonDropDown objects; 
    /// composition of IRibbonCommon, IDropDownItem &amp; IImageableItem</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IRibbonDropDown)]
    public interface IRibbonDropDown {
        /// <summary>Returns the unique (within this ribbon) identifier for this control.</summary>
        [DispId(DispIds.Id)]
        [Description("Returns the unique (within this ribbon) identifier for this control.")]
        string      Id                  { get; }
        /// <summary>Returns the Description string for this control. Only applicable for Menu Items.</summary>
        [Description("Returns the Description string for this control. Only applicable for Menu Items.")]
        [DispId(DispIds.Description)]
        string      Description         { get; }
        /// <summary>Returns the KeyTip string for this control.</summary>
        [Description("Returns the KeyTip string for this control.")]
        [DispId(DispIds.KeyTip)]
        string      KeyTip              { get; }
        /// <summary>Returns the Label string for this control.</summary>
        [Description("Returns the Label string for this control.")]
        [DispId(DispIds.Label)]
        string      Label               { get; }
        /// <summary>Returns the screenTip string for this control.</summary>
        [Description("Returns the screenTip string for this control.")]
        [DispId(DispIds.ScreenTip)]
        string      ScreenTip           { get; }
        /// <summary>Returns the SuperTip string for this control.</summary>
        [Description("Returns the SuperTip string for this control.")]
        [DispId(DispIds.SuperTip)]
        string      SuperTip            { get; }
        /// <summary>Sets the Label, KeyTip, ScreenTip and SuperTip for this control from the supplied values.</summary>
        [Description("Sets the Label, KeyTip, ScreenTip and SuperTip for this control from the supplied values.")]
        [DispId(DispIds.SetLanguageStrings)]
        void        SetLanguageStrings(IRibbonTextLanguageControl LanguageStrings);

        /// <summary>TODO</summary>
        [Description("")]
        [DispId(DispIds.IsEnabled)]
        bool        IsEnabled           { get; set; }
        /// <summary>TODO</summary>
        [Description("")]
        [DispId(DispIds.IsVisible)]
        bool        IsVisible           { get; set; }

        /// <summary>Returns the ID of the current selected item.</summary>
        [Description("Returns the ID of the current selected item.")]
        [DispId(DispIds.SelectedItemId)]
        string      SelectedItemId      { get; set; }
        /// <summary>Returns the index of the current selected item.</summary>
        [Description("Returns the index of the current selected item.")]
        [DispId(DispIds.SelectedItemIndex)]
        int         SelectedItemIndex   { get; set; }
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
