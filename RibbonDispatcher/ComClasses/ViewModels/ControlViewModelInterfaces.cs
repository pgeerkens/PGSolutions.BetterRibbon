////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    using VM = ComClasses.ViewModels;

    /// <summary>The base interface for <see cref="VM.AbstractControlVM{TSource}"/> implementations</summary>
    [CLSCompliant(true)]
    public interface IControlVM {
        /// <summary>Returns the unique (within this ribbon) identifier for this control.</summary>
        string Id           { get; }
        /// <summary>Returns the KeyTip string for this control.</summary>
        string KeyTip       { get; }
        /// <summary>Returns the Label string for this control.</summary>
        string Label        { get; }
        /// <summary>Gets or sets whether or not the control is visible.</summary>
        bool   IsVisible    { get; }

        /// <summary>Gets or sets whether or not the control is enabled.</summary>
        bool   IsEnabled    { get; }

        /// <summary>.</summary>
        void   Invalidate();

        /// <summary>.</summary>
        void   Detach();

        /// <summary>Returns the Description string for this control. Only applicable for Menu Items.</summary>
        string Description  { get; }
        /// <summary>Returns the screenTip string for this control.</summary>
        string ScreenTip    { get; }
        /// <summary>Returns the SuperTip string for this control.</summary>
        string SuperTip     { get; }
    }

    /// <summary>The total interface exposed by <see cref="VM.ButtonVM"/> objects.</summary>
    [CLSCompliant(true)]
    public interface IButtonVM: IControlVM, IClickableVM, IImageableVM, ISizeableVM { }

    /// <summary>The total interface exposed by <see cref="VM.CheckBoxVM"/> objects.</summary>
    [CLSCompliant(true)]
    public interface ICheckBoxVM: IControlVM, IToggleableVM { }

    /// <summary>The total interface exposed by <see cref="VM.CheckBoxVM"/> objects.</summary>
    [CLSCompliant(true)]
    public interface IToggleVM: IControlVM, IToggleableVM, IImageableVM, ISizeableVM { }

    /// <summary>The total interface exposed by <see cref="VM.EditBoxVM"/> objects.</summary>
    [CLSCompliant(true)]
    public interface IEditBoxVM : IControlVM, IEditableVM { }

    /// <summary>The total interface exposed by <see cref="VM.DropDownVM"/> objects.</summary>
    [CLSCompliant(true)]
    public interface IDropDownVM: IControlVM, ISelectableVM, ISelectable2VM { }

    /// <summary>The total interface exposed by <see cref="VM.SelectableItemVM"/> objects.</summary>
    [CLSCompliant(true)]
    public interface ISelectableItemVM: IControlVM, IImageableVM { }

    /// <summary>The total interface exposed by <see cref="VM.ComboBoxVM"/> objects.</summary>
    [CLSCompliant(true)]
    public interface IComboBoxVM: IControlVM, ISelectableVM, IEditBoxVM { }

    /// <summary>The total interface exposed by <see cref="VM.GroupVM"/> objects.</summary>
    [SuppressMessage("Microsoft.Design", "CA1040:AvoidEmptyInterfaces")]
    public interface IGroupVM: IControlVM { }

    /// <summary>The total interface exposed by <see cref="VM.TabVM"/> objects.</summary>
    [SuppressMessage("Microsoft.Design", "CA1040:AvoidEmptyInterfaces")]
    public interface ITabVM: IControlVM { }

    /// <summary>The total interface exposed by <see cref="VM.GroupVM"/> objects.</summary>
    [SuppressMessage("Microsoft.Design", "CA1040:AvoidEmptyInterfaces")]
    public interface IGalleryM: IControlVM, IGallerySizeVM, ISelectableVM, ISelectable2VM { }

    public interface ILabelVM: IControlVM { }

    public interface ISplitButtonVM: IControlVM {
        IMenuVM   MenuVM   { get; }
    }

    public interface ISplitToggleButtonVM: ISplitButtonVM {
        IToggleVM ToggleVM { get; }
    }

    public interface ISplitPressButtonVM: ISplitButtonVM {
        IButtonVM ButtonVM { get; }
    }

    public interface IMenuVM: IControlVM { }
}
