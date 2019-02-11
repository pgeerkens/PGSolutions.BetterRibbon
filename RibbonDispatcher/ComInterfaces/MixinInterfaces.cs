////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using stdole;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary>TODO</summary>
    [CLSCompliant( true )]
    public delegate void ClickedEventHandler(object sender);

    public delegate void ToggledEventHandler(object sender, bool isPressed);

    public delegate void SelectedEventHandler(object sender, int selectedIndex);

    public enum ERibbonControlSize {
        SizeSmall = RibbonControlSize.RibbonControlSizeRegular,
        SizeLarge = RibbonControlSize.RibbonControlSizeLarge
    }

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
        bool ShowImage { get; set; }
        /// <summary>Sets or gets whether to display the Label for this control.</summary>
        bool ShowLabel { get; set; }

        object Image { get; }

        /// <summary>Sets the displayable image for this control to the provided {IPictureDisp}</summary>
        void SetImageDisp(IPictureDisp Image);

        /// <summary>Sets the displayable image for this control to the named ImageMso image</summary>
        void SetImageMso(string ImageMso);
    }

    /// <summary>The interface for controls that can be sized.</summary>
    [CLSCompliant(true)]
    internal interface ISizeable {
        /// <summary>TODO</summary>
        void Invalidate();

        bool IsLarge { get; set; }
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

    public class EventArgs<T> : EventArgs where T:struct {
        public EventArgs(T value) : base() => Value = value;

        public T Value { get; }
    }

    /// <summary></summary>
    [Description("")]    
    [ComVisible(true)]
    [CLSCompliant(false)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IViewModelStore)]
    public interface IViewModelStore {
        IRibbonGroup    AttachGroup(string controlId, IRibbonControlStrings strings);
        IRibbonButton   AttachButton(string controlId, IRibbonControlStrings strings);
        IRibbonToggle   AttachToggle(string controlId, IRibbonControlStrings strings, IBooleanSource source);
        IRibbonToggle   AttachCheckBox(string controlId, IRibbonControlStrings strings, IBooleanSource source);
        IRibbonDropDown AttachDropDown(string controlId, IRibbonControlStrings strings, IIntegerSource source);
    }

    /// <summary></summary>
    [Description("")]    
    [ComVisible(true)]
    [CLSCompliant(false)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IRibbonGroupModel)]
    public interface IRibbonGroupModel {
        IRibbonGroup ViewModel { get; }

        void Attach(string controlId, IRibbonControlStrings strings);
    }

    /// <summary></summary>
    [Description("")]    
    [ComVisible(true)]
    [CLSCompliant(false)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IRibbonButtonModel)]
    public interface IRibbonButtonModel {
        [SuppressMessage( "Microsoft.Design", "CA1009:DeclareEventHandlersCorrectly", Justification="EventArgs<T> is unknown to COM.")]
        event ClickedEventHandler Clicked;

        IRibbonButton ViewModel { get; }

        void Attach(string controlId, IRibbonControlStrings strings);
    }

    /// <summary></summary>
    [Description("")]    
    [ComVisible(true)]
    [CLSCompliant(false)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IRibbonToggleModel)]
    public interface IRibbonToggleModel {
        [SuppressMessage( "Microsoft.Design", "CA1009:DeclareEventHandlersCorrectly", Justification="EventArgs<T> is unknown to COM.")]
        event ToggledEventHandler Toggled;

        IRibbonToggle ViewModel { get; }
        bool          IsPressed { get; set; }

        IRibbonToggleModel Attach(string controlId, IRibbonControlStrings strings);
    }

    public static partial class Extensions {
        public static RibbonControlSize ControlSize(this bool isLarge)
            => isLarge ? RibbonControlSize.RibbonControlSizeLarge
                       : RibbonControlSize.RibbonControlSizeRegular;
    }
}
