////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using Microsoft.Office.Core;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary>The interface for controls that can have images.</summary>
    [CLSCompliant(true)]
    internal interface IImageable {
        /// <summary>TODO</summary>
        void Invalidate();

        /// <summary>Sets or gets whether to display the Image for this control.</summary>
        bool    ShowImage         { get; }

        /// <summary>Sets or gets whether to display the Label for this control.</summary>
        bool    ShowLabel         { get; }

        ImageObject Image         { get; }
    }

    /// <summary>The interface for controls that can be sized.</summary>
    [CLSCompliant(true)]
    internal interface ISizeable {
        /// <summary>TODO</summary>
        void Invalidate();

        bool    IsLarge           { get; }
    }

    /// <summary>The interface for controls that can be clicked.</summary>
    [CLSCompliant(true)]
    internal interface IClickable {
        /// <summary>TODO</summary>
        void OnClicked(IRibbonControl control);
    }

    /// <summary>The interface for controls that can be toggled.</summary>
    [CLSCompliant(true)]
    internal interface IToggleable {
        /// <summary>TODO</summary>
        void Invalidate();

        /// <summary>TODO</summary>
        void OnToggled(IRibbonControl control, bool isPressed);

        bool    IsPressed         { get; }
    }

    public interface IEditable {
        /// <summary>Current contents of this <see cref="ITextEditable"/> control.</summary>
        string  Text              { get; }

        /// <summary>Call back for OnChanged events from <see cref="ITextEditable"/> controls.</summary>
        void OnEdited(IRibbonControl sender, string text);
    }

    /// <summary>The interface for controls that have a selectable list of items.</summary>
    [CLSCompliant(true)]
    internal interface ISelectable {
        /// <summary>ID of the selected item.</summary>
        string  SelectedItemId    { get; }
        /// <summary>Index of the selected item.</summary>
        int     SelectedItemIndex { get; }
        /// <summary>Call back for OnAction events from the drop-down ribbon elements.</summary>
        void    OnSelected(IRibbonControl control, string selectedId, int selectedIndex);

        /// <summary>Call back for ItemCount events from the drop-down ribbon elements.</summary>
        int     ItemCount         { get; }
        /// <summary>Call back for GetItemID events from the drop-down ribbon elements.</summary>
        string  ItemId(int Index);
        /// <summary>Call back for GetItemLabel events from the drop-down ribbon elements.</summary>
        string  ItemLabel(int Index);
        /// <summary>Call back for GetItemScreenTip events from the drop-down ribbon elements.</summary>
        string  ItemScreenTip(int Index);
        /// <summary>Call back for GetItemSuperTip events from the drop-down ribbon elements.</summary>
        string  ItemSuperTip(int Index);

        /// <summary>Call back for GetItemSuperTip events from the drop-down ribbon elements.</summary>
        object  ItemImage(int Index);
        /// <summary>Call back for GetItemSuperTip events from the drop-down ribbon elements.</summary>
        bool    ItemShowImage(int Index);
        /// <summary>Call back for GetItemSuperTip events from the drop-down ribbon elements.</summary>
        bool    ItemShowLabel(int Index);
    }
}
