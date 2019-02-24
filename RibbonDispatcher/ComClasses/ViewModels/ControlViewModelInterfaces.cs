////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary>The base interface for Ribbon controls.</summary>
    [CLSCompliant(true)]
    public interface IControlVM {
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
    public interface IButtonVM: IControlVM, IImageableVM, ISizeableVM, IClickableVM { }

    public interface ICheckBoxVM: IToggleableVM, IControlVM { }

    /// <summary>The ViewModel interface exposed by Ribbon ToggleButtons and CheckBoxes.</summary>
    public interface IToggleControlVM: IToggleableVM, IControlVM, IImageableVM, ISizeableVM { }

    public interface IEditBoxVM : IEditableVM, IControlVM { }

    /// <summary>The total interface (required to be) exposed externally by DropDownVM objects; 
    /// composition of IControlVM, IDropDownItem &amp; IImageableItem</summary>
    public interface IDropDownVM: IControlVM, ISelectableVM, ISelectable2VM { }

    public interface ISelectableItemVM: IControlVM, IImageableVM { }

    public interface IComboBoxVM: IControlVM, ISelectableVM, IEditBoxVM { }

    /// <summary>TODO</summary>
    [SuppressMessage("Microsoft.Design", "CA1040:AvoidEmptyInterfaces")]
    public interface IGroupVM: IControlVM { }
}
