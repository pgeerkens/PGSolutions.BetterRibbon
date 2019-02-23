////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

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
    public interface IButtonVM: IRibbonControlVM, IImageableVM, ISizeableVM, IClickableVM { }

    public interface ICheckBoxVM: IToggleableVM, IRibbonControlVM { }

    /// <summary>The ViewModel interface exposed by Ribbon ToggleButtons and CheckBoxes.</summary>
    public interface IToggleControlVM: IToggleableVM, IRibbonControlVM, IImageableVM, ISizeableVM { }

    public interface IEditBoxVM : IEditableVM, IRibbonControlVM { }

    /// <summary>The total interface (required to be) exposed externally by DropDownVM objects; 
    /// composition of IRibbonControlVM, IDropDownItem &amp; IImageableItem</summary>
    public interface IDropDownVM: IRibbonControlVM, ISelectableVM { }

    public interface IComboBoxVM: IDropDownVM, IEditBoxVM { }

    /// <summary>TODO</summary>
    public interface IGroupVM: IRibbonControlVM { }
}
