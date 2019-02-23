////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

using Microsoft.Office.Core;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary>The base interface for Ribbon controls.</summary>
    [CLSCompliant(true)]
    public interface IRibbonControlVM {
        /// <summary>Returns the unique (within this ribbon) identifier for this control.</summary>
        string Id { get; }
        /// <summary>Returns the Description string for this control. Only applicable for Menu Items.</summary>
        string Description { get; }
        /// <summary>Returns the KeyTip string for this control.</summary>
        string KeyTip { get; }
        /// <summary>Returns the Label string for this control.</summary>
        string Label { get; }
        /// <summary>Returns the screenTip string for this control.</summary>
        string ScreenTip { get; }
        /// <summary>Returns the SuperTip string for this control.</summary>
        string SuperTip { get; }

        /// <summary>Gets or sets whether or not the control is enabled.</summary>
        bool IsEnabled { get; }
        /// <summary>Gets or sets whether or not the control is visible.</summary>
        bool IsVisible { get; }

        void Invalidate();

        void Detach();
    }

    /// <summary>The total interface (required to be) exposed externally by ButtonVM objects.</summary>
    public interface IButtonVM: IRibbonControlVM, IImageableVM, ISizeableVM, IClickableVM {
        ///// <summary>Returns the unique (within this ribbon) identifier for this control.</summary>
        //[Description("Returns the unique (within this ribbon) identifier for this control.")]
        //new string Id { get; }
        ///// <summary>Returns the Description string for this control. Only applicable for Menu Items.</summary>
        //[Description("Returns the Description string for this control. Only applicable for Menu Items.")]
        //new string Description { get; }
        ///// <summary>Returns the KeyTip string for this control.</summary>
        //[Description("Returns the KeyTip string for this control.")]
        //new string KeyTip { get; }
        ///// <summary>Returns the Label string for this control.</summary>
        //[Description("Returns the Label string for this control.")]
        //new string Label { get; }
        ///// <summary>Returns the screenTip string for this control.</summary>
        //[Description("Returns the screenTip string for this control.")]
        //new string ScreenTip { get; }
        ///// <summary>Returns the SuperTip string for this control.</summary>
        //[Description("Returns the SuperTip string for this control.")]
        //new string SuperTip { get; }

        ///// <summary>Gets or sets whether the control is enabled.</summary>
        //[Description("Gets or sets whether the control is enabled.")]
        //new bool IsEnabled { get; }
        ///// <summary>Gets or sets whether the control is visible.</summary>
        //[Description("Gets or sets whether the control is visible.")]
        //new bool IsVisible { get; }

        ///// <inheritdoc/>
        //new void Invalidate();

        ///// <summary>Gets or sets the preferred {RibbonControlSize} for the control.</summary>
        //[Description("Gets or sets the preferred {RdControlSize} for the control.")]
        //bool IsLarge { get; }

        ///// <summary>Callback for the Clicked event on the control.</summary>
        //[Description("Callback for the Clicked event on the control.")]
        //void OnClicked(IRibbonControl control);

        ///// <summary>Returns the current Image for the control as either a {string} naming an MsoImage or an {IPictureDisp}.</summary>
        //[Description("Returns the current Image for the control as either a {string} naming an MsoImage or an {IPictureDisp}.")]
        //new ImageObject Image { get; }
        ///// <summary>Gets or sets whether to show the control's image; ignored by Large controls.</summary>
        //[Description("Gets or sets whether to show the control's image; ignored by Large controls.")]
        //new bool ShowImage { get; }
        ///// <summary>Gets or sets whether to show the control's label; ignored by Large controls.</summary>
        //[Description("Gets or sets whether to show the control's label; ignored by Large controls.")]
        //new bool ShowLabel { get; }
    }

    /// <summary>The ViewModel interface exposed by Ribbon ToggleButtons and CheckBoxes.</summary>
    public interface IToggleButtonVM: IToggleableVM, IRibbonControlVM, IImageableVM, ISizeableVM {
        ///// <summary>Returns the unique (within this ribbon) identifier for this control.</summary>
        //new string Id { get; }

        ///// <summary>Returns the Description string for this control. Only applicable for Menu Items.</summary>
        //[Description("Returns the Description string for this control. Only applicable for Menu Items.")]
        //new string Description { get; }

        ///// <summary>Returns the KeyTip string for this control.</summary>
        //[Description("Returns the KeyTip string for this control.")]
        //new string KeyTip { get; }

        ///// <summary>Returns the Label string for this control.</summary>
        //[Description("Returns the Label string for this control.")]
        //new string Label { get; }

        ///// <summary>Returns the screenTip string for this control.</summary>
        //[Description("Returns the screenTip string for this control.")]
        //new string ScreenTip { get; }

        ///// <summary>Returns the SuperTip string for this control.</summary>
        //[Description("Returns the SuperTip string for this control.")]
        //new string SuperTip { get; }

        ///// <summary>Gets or sets whether the control is enabled.</summary>
        //[Description("Gets or sets whether the control is enabled.")]
        //new bool IsEnabled { get; }
        ///// <summary>Gets or sets whether the control is visible.</summary>
        //[Description("Gets or sets whether the control is visible.")]
        //new bool IsVisible { get; }

        ///// <inheritdoc/>
        //new void Invalidate();

        ///// <summary>Returns whether the control is pressed.</summary>
        //[Description("Returns whether the control is pressed.")]
        //bool IsPressed { get; }
        ///// <summary>Callback for the Pressed event on the control.</summary>
        //[Description("Callback for the Pressed event on the control.")]
        //void OnToggled(IRibbonControl control, bool isPressed);

        ///// <summary>TODO</summary>
        //bool IsSizeable { get; }
        ///// <summary>TODO</summary>
        //bool IsLarge { get; }

        ///// <summary>TODO</summary>
        //bool IsImageable { get; }
        ///// <summary>Gets or sets whether to show the control's image; ignored by Large controls.</summary>
        //[Description("Gets or sets whether to show the control's image; ignored by Large controls.")]
        //new bool ShowImage { get; }
        ///// <summary>Gets or sets whether to show the control's label; ignored by Large controls.</summary>
        //[Description("Gets or sets whether to show the control's label; ignored by Large controls.")]
        //new bool ShowLabel { get; }
        ///// <summary>Returns the current Image for the control as either a {string} naming an MsoImage or an {IPictureDisp}.</summary>
        //[Description("Returns the current Image for the control as either a {string} naming an MsoImage or an {IPictureDisp}.")]
        //new ImageObject Image { get; }
    }

    public interface IEditBoxVM : IEditableVM, IRibbonControlVM {
        ///// <summary>Returns the unique (within this ribbon) identifier for this control.</summary>
        //[Description("Returns the unique (within this ribbon) identifier for this control.")]
        //new string Id           { get; }
        ///// <summary>Returns the Description string for this control. Only applicable for Menu Items.</summary>
        //[Description("Returns the Description string for this control. Only applicable for Menu Items.")]
        //new string Description  { get; }
        ///// <summary>Returns the KeyTip string for this control.</summary>
        //[Description("Returns the KeyTip string for this control.")]
        //new string KeyTip       { get; }
        ///// <summary>Returns the Label string for this control.</summary>
        //[Description("Returns the Label string for this control.")]
        //new string Label        { get; }
        ///// <summary>Returns the screenTip string for this control.</summary>
        //[Description("Returns the screenTip string for this control.")]
        //new string ScreenTip    { get; }
        ///// <summary>Returns the SuperTip string for this control.</summary>
        //[Description("Returns the SuperTip string for this control.")]
        //new string SuperTip     { get; }

        ///// <summary>Gets or sets whether the control is enabled.</summary>
        //[Description("Gets or sets whether the control is enabled.")]
        //new bool IsEnabled      { get; }
        ///// <summary>Gets or sets whether the control is visible.</summary>
        //[Description("Gets or sets whether the control is visible.")]
        //new bool IsVisible      { get; }

        ///// <inheritdoc/>
        //new void Invalidate();
    }
    /// <summary>The total interface (required to be) exposed externally by DropDownVM objects; 
    /// composition of IRibbonControlVM, IDropDownItem &amp; IImageableItem</summary>
    public interface IDropDownVM: IRibbonControlVM {
        ///// <summary>Returns the unique (within this ribbon) identifier for this control.</summary>
        //[Description("Returns the unique (within this ribbon) identifier for this control.")]
        //new string Id { get; }
        ///// <summary>Returns the Description string for this control. Only applicable for Menu Items.</summary>
        //[Description("Returns the Description string for this control. Only applicable for Menu Items.")]
        //new string Description { get; }
        ///// <summary>Returns the KeyTip string for this control.</summary>
        //[Description("Returns the KeyTip string for this control.")]
        //new string KeyTip { get; }
        ///// <summary>Returns the Label string for this control.</summary>
        //[Description("Returns the Label string for this control.")]
        //new string Label { get; }
        ///// <summary>Returns the screenTip string for this control.</summary>
        //[Description("Returns the screenTip string for this control.")]
        //new string ScreenTip { get; }
        ///// <summary>Returns the SuperTip string for this control.</summary>
        //[Description("Returns the SuperTip string for this control.")]
        //new string SuperTip { get; }

        ///// <summary>TODO</summary>
        //[Description("")]
        //new bool IsEnabled { get; }
        ///// <summary>TODO</summary>
        //[Description("")]
        //new bool IsVisible { get; }

        ///// <inheritdoc/>
        //new void Invalidate();

        /// <summary>Returns the ID of the current selected item.</summary>
        string SelectedItemId { get; }
        /// <summary>Returns the index of the current selected item.</summary>
        int SelectedItemIndex { get; }
        /// <summary>Call back for OnAction events from the drop-down ribbon elements.</summary>
        void OnSelectionMade(IRibbonControl control, string selectedId, int selectedIndex);

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
        bool ItemShowImage(int Index);
        /// <summary>Call back for GetItemSuperTip events from the drop-down ribbon elements.</summary>
        bool ItemShowLabel(int Index);
    }

    public interface IComboBoxVM: IDropDownVM, IEditBoxVM {

    }

    /// <summary>TODO</summary>
    public interface IGroupVM: IRibbonControlVM {
        ///// <summary>Returns the unique (within this ribbon) identifier for this control.</summary>
        //[Description("Returns the unique (within this ribbon) identifier for this control.")]
        //new string Id           { get; }

        ///// <summary>Returns the Description string for this control. Only applicable for Menu Items.</summary>
        //[Description("Returns the Description string for this control. Only applicable for Menu Items.")]
        //new string Description  { get; }

        ///// <summary>Returns the KeyTip string for this control.</summary>
        //[Description("Returns the KeyTip string for this control.")]
        //new string KeyTip       { get; }

        ///// <summary>Returns the Label string for this control.</summary>
        //[Description("Returns the Label string for this control.")]
        //new string Label        { get; }

        ///// <summary>Returns the screenTip string for this control.</summary>
        //[Description("Returns the screenTip string for this control.")]
        //new string ScreenTip    { get; }

        ///// <summary>Returns the SuperTip string for this control.</summary>
        //[Description("Returns the SuperTip string for this control.")]
        //new string SuperTip     { get; }

        ///// <summary>Gets or sets whether the control is enabled.</summary>
        //new bool IsEnabled {
        //    [Description("Gets or sets whether the control is enabled.")]
        //    get; }

        ///// <summary>Gets or sets whether the control is visible.</summary>
        //new bool IsVisible {
        //    [Description("Gets or sets whether the control is visible.")]
        //    get; }

        ///// <summary>.</summary>
        //[Description(".")]
        //new void Invalidate();
    }
}
