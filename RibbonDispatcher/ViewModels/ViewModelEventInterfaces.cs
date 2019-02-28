////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;

using Microsoft.Office.Core;

namespace PGSolutions.RibbonDispatcher.ViewModels {
    public delegate void ClickedEventHandler(IRibbonControl control);

    public delegate void ToggledEventHandler(IRibbonControl control, bool isPressed);

    public delegate void SelectionMadeEventHandler(IRibbonControl control, string selectedId, int selectedIndex);

    public delegate void EditedEventHandler(IRibbonControl control, string text);

    /// <summary>.</summary>
    /// <typeparam name="T"></typeparam>
    public class EventArgs<T>:EventArgs {
        public EventArgs(T value) : base() => Value = value;

        public T Value { get; }
    }

    public class ToggledEventArgs : EventArgs<bool> {
        public ToggledEventArgs(bool isPressed) : base(isPressed) { }
    }

    public class SelectedEventArgs : EventArgs<ValueTuple<string, int>> {
        public SelectedEventArgs(string SelectedId, int SelectedIndex)
        : base((SelectedId, SelectedIndex)) { }
    }

    public class EditedEventArgs : EventArgs<string> {
        EditedEventArgs(string text) : base(text) { }
    }

    public delegate void ClickedEventHandler2(object sender, EventArgs e);

    public delegate void ToggledEventHandler2(object sender, ToggledEventArgs e);

    public delegate void SelectionMadeEventHandler2(object sender, SelectedEventArgs e);

    public delegate void EditedEventHandler2(object sender, EditedEventArgs e);

    /// <summary>The interface for controls that can have images.</summary>
    [CLSCompliant(true)]
    public interface IImageableVM {
        /// <summary>Returns the current Image for the control as either a {string} naming an MsoImage or an {IPictureDisp}.</summary>
        ImageObject Image       { get; }
        /// <summary>Gets or sets whether to show the control's image; ignored by Large controls.</summary>
        bool        ShowImage   { get; }
        /// <summary>Gets or sets whether to show the control's label; ignored by Large controls.</summary>
        bool        ShowLabel   { get; }
    }

    /// <summary>The interface for controls that can be sized.</summary>
    [CLSCompliant(true)]
    public interface ISizeableVM {
        bool    IsLarge           { get; }
    }

    /// <summary>The interface for controls that can be clicked.</summary>
    [CLSCompliant(true)]
    public interface IClickableVM {
        /// <summary>TODO</summary>
        [SuppressMessage("Microsoft.Design", "CA1009:DeclareEventHandlersCorrectly")]
        event ClickedEventHandler Clicked;

        /// <summary>TODO</summary>
        void OnClicked(IRibbonControl control);
    }

    /// <summary>The interface for controls that can be toggled.</summary>
    [CLSCompliant(true)]
    public interface IToggleableVM {
        /// <summary>TODO</summary>
        [SuppressMessage("Microsoft.Design", "CA1009:DeclareEventHandlersCorrectly")]
        event ToggledEventHandler Toggled;

        /// <summary>TODO</summary>
        void OnToggled(IRibbonControl control, bool isPressed);

        bool    IsPressed         { get; }
    }

    [CLSCompliant(true)]
    public interface IEditableVM {
        /// <summary>TODO</summary>
        [SuppressMessage("Microsoft.Design", "CA1009:DeclareEventHandlersCorrectly")]
        event EditedEventHandler Edited;

        /// <summary>Current contents of this <see cref="ITextEditable"/> control.</summary>
        string  Text              { get; }

        /// <summary>Call back for OnChanged events from <see cref="ITextEditable"/> controls.</summary>
        void OnEdited(IRibbonControl sender, string text);
    }

    [CLSCompliant(true)]
    public interface IStaticListVM {
        IReadOnlyList<IStaticItemVM> Items { get; }
    }

    /// <summary>The interface for controls that have a selectable list of items.</summary>
    [CLSCompliant(true)]
    public interface ISelectItemsVM {    // DropDown & ComboBox & Gallery
        ///// <summary>Call back for ItemCount events from the drop-down ribbon elements.</summary>
        //int ItemCount { get; }

        //IStaticItemVM this[int index] { get; }

        IReadOnlyList<IStaticItemVM> Items { get; }
    }

    /// <summary>The interface for controls that have a selectable list of items.</summary>
    [CLSCompliant(true)]
    public interface ISelectablesVM {   // DropDown
        [SuppressMessage("Microsoft.Design", "CA1009:DeclareEventHandlersCorrectly")]
        event SelectionMadeEventHandler SelectionMade;

        /// <summary>ID of the selected item.</summary>
        string  SelectedItemId    { get; }
        /// <summary>Index of the selected item.</summary>
        int     SelectedItemIndex { get; }

        /// <summary>Call back for OnAction events from the drop-down ribbon elements.</summary>
        void OnSelectionMade(IRibbonControl control, string selectedId, int selectedIndex);
    }

    /// <summary>The interface for controls that have a selectable list of items.</summary>
    [CLSCompliant(true)]
    public interface IDynamicMenuVM {
        string MenuContent { get; }
    }

    /// <summary>The interface for galleries with sizeable items.</summary>
    [CLSCompliant(true)]
    public interface IGallerySizeVM {
        int ItemHeight { get; }

        int ItemWidth  { get; }
    }

    [CLSCompliant(true)]
    public interface IDescriptionableVM {
        string Description { get; }
    }

    [CLSCompliant(true)]
    public interface IMenuSeparatorVM {
        string Title { get; }
    }
}
