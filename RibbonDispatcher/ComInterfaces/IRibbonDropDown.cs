////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.ComponentModel;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary>The total interface (required to be) exposed externally by DropDownVM objects; 
    /// composition of IRibbonControlVM, IDropDownItem &amp; IImageableItem</summary>
    public interface IRibbonDropDown : IRibbonControlVM {
        /// <summary>Returns the unique (within this ribbon) identifier for this control.</summary>
        [Description("Returns the unique (within this ribbon) identifier for this control.")]
        new string      Id         { get; }
        /// <summary>Returns the Description string for this control. Only applicable for Menu Items.</summary>
        [Description("Returns the Description string for this control. Only applicable for Menu Items.")]
        new string Description     { get; }
        /// <summary>Returns the KeyTip string for this control.</summary>
        [Description("Returns the KeyTip string for this control.")]
        new string KeyTip          { get; }
        /// <summary>Returns the Label string for this control.</summary>
        [Description("Returns the Label string for this control.")]
        new string Label           { get; }
        /// <summary>Returns the screenTip string for this control.</summary>
        [Description("Returns the screenTip string for this control.")]
        new string ScreenTip       { get; }
        /// <summary>Returns the SuperTip string for this control.</summary>
        [Description("Returns the SuperTip string for this control.")]
        new string SuperTip        { get; }

        /// <summary>TODO</summary>
        [Description("")]
        new bool IsEnabled       { get; }
        /// <summary>TODO</summary>
        [Description("")]
        new bool IsVisible       { get; }

        /// <inheritdoc/>
        new void Invalidate();

        /// <summary>Returns the ID of the current selected item.</summary>
        [Description("Returns the ID of the current selected item.")]
        string      SelectedItemId      { get; }
        /// <summary>Returns the index of the current selected item.</summary>
        [Description("Returns the index of the current selected item.")]
        int         SelectedItemIndex   { get; }
        /// <summary>Call back for OnAction events from the drop-down ribbon elements.</summary>
        [Description("Call back for OnAction events from the drop-down ribbon elements.")]
        void        OnActionDropDown(string SelectedId, int SelectedIndex);

        /// <summary>Call back for ItemCount events from the drop-down ribbon elements.</summary>
        [Description("Call back for ItemCount events from the drop-down ribbon elements.")]
        int         ItemCount           { get; }
        /// <summary>Call back for GetItemID events from the drop-down ribbon elements.</summary>
        [Description("Call back for GetItemID events from the drop-down ribbon elements.")]
        string      ItemId(int Index);
        /// <summary>Call back for GetItemLabel events from the drop-down ribbon elements.</summary>
        [Description("Call back for GetItemLabel events from the drop-down ribbon elements")]
        string      ItemLabel(int Index);
        /// <summary>Call back for GetItemScreenTip events from the drop-down ribbon elements.</summary>
        [Description("Call back for GetItemScreenTip events from the drop-down ribbon elements.")]
        string      ItemScreenTip(int Index);
        /// <summary>Call back for GetItemSuperTip events from the drop-down ribbon elements.</summary>
        [Description("Call back for GetItemSuperTip events from the drop-down ribbon elements.")]
        string      ItemSuperTip(int Index);

        /// <summary>Call back for GetItemSuperTip events from the drop-down ribbon elements.</summary>
        [Description("Call back for GetItemSuperTip events from the drop-down ribbon elements.")]
        bool        ItemShowImage(int Index);
        /// <summary>Call back for GetItemSuperTip events from the drop-down ribbon elements.</summary>
        [Description("Call back for GetItemSuperTip events from the drop-down ribbon elements.")]
        bool        ItemShowLabel(int Index);
    }
}
