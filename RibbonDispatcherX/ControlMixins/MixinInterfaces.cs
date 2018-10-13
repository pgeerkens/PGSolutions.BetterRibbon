////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.InteropServices;
using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ControlMixins {
    /// <summary>The interface for controls that are Clickable.</summary>
    [CLSCompliant(true)]
    public interface IClickableMixin {
        /// <summary>TODO</summary>
        void OnClicked();
    }

    /// <summary>The interface for controls that are Imageable.</summary>
    [CLSCompliant(true)]
    public interface IImageableMixin {
        /// <summary>TODO</summary>
        void OnChanged();
    }

    /// <summary>The interface for controls that can be sized.</summary>
    [CLSCompliant(true)]
    public interface ISizeableMixin {
        /// <summary>TODO</summary>
        void OnChanged();
    }

    /// <summary>The interface for controls that can be toggled.</summary>
    [CLSCompliant(true)]
    internal interface IToggleableMixin {
        /// <summary>TODO</summary>
        void OnChanged();

        /// <summary>TODO</summary>
        void OnToggled(bool IsPressed);

        /// <summary>TODO</summary>
        IRibbonControlStrings LanguageStrings { get; }
    }

    /// <summary>TODO</summary>
    [CLSCompliant(true)]
    public delegate void SelectedEventHandler(string ItemId, int ItemIndex);

    /// <summary>TODO</summary>
    [CLSCompliant(true)]
    public interface ISelectableMixin {
        /// <summary>ID of the selected item.</summary>
        [DispId(DispIds.SelectedItemId)]
        string SelectedItemId { get; }
        /// <summary>Index of the selected item.</summary>
        [DispId(DispIds.SelectedItemIndex)]
        int SelectedItemIndex { get; }
        /// <summary>Call back for OnAction events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.OnActionDropDown)]
        void OnActionDropDown(string SelectedId, int SelectedIndex);

        /// <summary>Call back for ItemCount events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.ItemCount)]
        int ItemCount { get; }
        /// <summary>Call back for GetItemID events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.ItemId)]
        string ItemId(int Index);
        /// <summary>Call back for GetItemLabel events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.ItemLabel)]
        string ItemLabel(int Index);
        /// <summary>Call back for GetItemScreenTip events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.ItemScreenTip)]
        string ItemScreenTip(int Index);
        /// <summary>Call back for GetItemSuperTip events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.ItemSuperTip)]
        string ItemSuperTip(int Index);

        /// <summary>Call back for GetItemSuperTip events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.ItemImage)]
        object ItemImage(int Index);
        /// <summary>Call back for GetItemSuperTip events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.ItemShowImage)]
        bool ItemShowImage(int Index);
        /// <summary>Call back for GetItemSuperTip events from the drop-down ribbon elements.</summary>
        [DispId(DispIds.ItemShowLabel)]
        bool ItemShowLabel(int Index);
    }
}
