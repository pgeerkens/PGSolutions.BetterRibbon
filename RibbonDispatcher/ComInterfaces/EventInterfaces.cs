////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    public enum ERibbonControlSize {
        SizeSmall = RibbonControlSize.RibbonControlSizeRegular,
        SizeLarge = RibbonControlSize.RibbonControlSizeLarge
    }

    public class EventArgs<T>:EventArgs {
        public EventArgs(T value) : base() => Value = value;

        public T Value { get; }
    }

    /// <summary>TODO</summary>
    [CLSCompliant( true )]
    public delegate void ClickedEventHandler(object sender);

    public delegate void ToggledEventHandler(object sender, bool isPressed);

    public delegate void SelectedEventHandler(object sender, int selectedIndex);

    /// <summary>The interface for controls that can be clicked.</summary>
    [CLSCompliant(true)]
    internal interface IClickable {
        /// <summary>TODO</summary>
        void OnClicked(object sender);
    }

    /// <summary>The interface for controls that can have images.</summary>
    [CLSCompliant(true)]
    internal interface IImageable {
        /// <summary>TODO</summary>
        void Invalidate();

        /// <summary>Sets or gets whether to display the Image for this control.</summary>
        bool ShowImage    { get; }
        /// <summary>Sets or gets whether to display the Label for this control.</summary>
        bool ShowLabel    { get; }

        ImageObject Image { get; }
    }

    /// <summary>The interface for controls that can be sized.</summary>
    [CLSCompliant(true)]
    internal interface ISizeable {
        /// <summary>TODO</summary>
        void Invalidate();

        bool IsLarge { get; }
    }

    /// <summary>The interface for controls that can be toggled.</summary>
    [CLSCompliant(true)]
    internal interface IToggleable {
        /// <summary>TODO</summary>
        void Invalidate();

        /// <summary>TODO</summary>
        void OnToggled(object sender, bool isPressed);

        bool IsPressed { get; }
    }

    /// <summary>The interface for controls that have a selectable list of items.</summary>
    [CLSCompliant(true)]
    internal interface ISelectable {
        /// <summary>ID of the selected item.</summary>
        string SelectedItemId { get; }
        /// <summary>Index of the selected item.</summary>
        int SelectedItemIndex { get; }
        /// <summary>Call back for OnAction events from the drop-down ribbon elements.</summary>
        void OnActionDropDown(string SelectedId, int SelectedIndex);

        /// <summary>Call back for ItemCount events from the drop-down ribbon elements.</summary>
        int ItemCount { get; }
        /// <summary>Call back for GetItemID events from the drop-down ribbon elements.</summary>
        string ItemId(int Index);
        /// <summary>Call back for GetItemLabel events from the drop-down ribbon elements.</summary>
        string ItemLabel(int Index);
        /// <summary>Call back for GetItemScreenTip events from the drop-down ribbon elements.</summary>
        string ItemScreenTip(int Index);
        /// <summary>Call back for GetItemSuperTip events from the drop-down ribbon elements.</summary>
        string ItemSuperTip(int Index);

        /// <summary>Call back for GetItemSuperTip events from the drop-down ribbon elements.</summary>
        object ItemImage(int Index);
        /// <summary>Call back for GetItemSuperTip events from the drop-down ribbon elements.</summary>
        bool ItemShowImage(int Index);
        /// <summary>Call back for GetItemSuperTip events from the drop-down ribbon elements.</summary>
        bool ItemShowLabel(int Index);
    }
}
