////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

namespace PGSolutions.RibbonDispatcher.ViewModels {
    /// <summary>The base interface for <see cref="AbstractControlVM{TSource}"/> implementations</summary>
    [CLSCompliant(true)]
    public interface IControlVM {
        /// <summary>Returns the unique (within this ribbon) identifier for this control.</summary>
        string ControlId    { get; }
        /// <summary>Gets or sets whether or not the control is visible.</summary>
        bool   IsVisible    { get; }
        /// <summary>Gets or sets whether or not the control is enabled.</summary>
        bool   IsEnabled    { get; }

        /// <summary>Returns the Label string for this control.</summary>
        string Label        { get; }
        /// <summary>Returns the screenTip string for this control.</summary>
        string ScreenTip    { get; }
        /// <summary>Returns the SuperTip string for this control.</summary>
        string SuperTip     { get; }
        /// <summary>Returns the KeyTip string for this control.</summary>
        string KeyTip       { get; }

        /// <summary>.</summary>
        void   Invalidate();

        /// <summary>.</summary>
        void   Detach();

        void OnPurged(IContainerControl sender);
        void SetShowInactive(bool showInactive);
    }

    /// <summary>The total interface exposed by <see cref="ButtonVM"/> objects.</summary>
    [CLSCompliant(true)]
    public interface IButtonVM: IControlVM, IClickableVM, IImageableVM, ISizeableVM, IDescriptionableVM { }

    /// <summary>The total interface exposed by <see cref="CheckBoxVM"/> objects.</summary>
    [CLSCompliant(true)]
    public interface ICheckBoxVM: IControlVM, IToggleableVM, IDescriptionableVM { }

    /// <summary>The total interface exposed by <see cref="ToggleButtonVM"/> objects.</summary>
    [CLSCompliant(true)]
    public interface IToggleVM: IControlVM, IToggleableVM, IImageableVM, ISizeableVM, IDescriptionableVM { }

    /// <summary>The total interface exposed by <see cref="EditBoxVM"/> objects.</summary>
    [CLSCompliant(true)]
    public interface IEditBoxVM : IControlVM, IEditableVM { }

    /// <summary>The total interface exposed by <see cref="DropDownVM"/> objects.</summary>
    [CLSCompliant(true)]
    public interface IDropDownVM: IControlVM, ISelectItemsVM, ISelectablesVM { }

    /// <summary>The total interface exposed by <see cref="GalleryVM"/> objects.</summary>
    [CLSCompliant(true)]
    public interface IGalleryVM: IControlVM, IGallerySizeVM, ISelectItemsVM, ISelectablesVM, IDescriptionableVM { }

    /// <summary>The total interface exposed by <see cref="GalleryVM"/> objects.</summary>
    [CLSCompliant(true)]
    public interface IStaticGalleryVM: IControlVM, IStaticListVM, IGallerySizeVM, ISelectablesVM, IDescriptionableVM { }

    /// <summary>The total interface exposed by <see cref="StaticItemVM"/> objects.</summary>
    [CLSCompliant(true)]
    public interface IStaticItemVM: IControlVM, IImageableVM { }

    /// <summary>The total interface exposed by <see cref="ComboBoxVM"/> objects.</summary>
    [CLSCompliant(true)]
    public interface IComboBoxVM: IControlVM, ISelectItemsVM, IEditBoxVM { }

    /// <summary>The total interface exposed by <see cref="ComboBoxVM"/> objects.</summary>
    [CLSCompliant(true)]
    public interface IStaticComboBoxVM: IControlVM, IStaticListVM, IEditBoxVM { }

    /// <summary>The total interface exposed by <see cref="GroupVM"/> objects.</summary>
    [CLSCompliant(true)]
    public interface IGroupVM: IControlVM {
        void Invalidate(Action<IControlVM> action);
    }

    /// <summary>The total interface exposed by <see cref="TabVM"/> objects.</summary>
    [CLSCompliant(true)]
    public interface ITabVM: IControlVM { }

    /// <summary>The total interface exposed by <see cref="LabelControlVM"/> objects.</summary>
    [CLSCompliant(true)]
    public interface ILabelControlVM: IControlVM { }

    /// <summary>The total interface exposed by <see cref="BoxControlVM"/> objects.</summary>
    [CLSCompliant(true)]
    public interface IBoxControlVM: IControlVM { }

    /// <summary>The total interface exposed by <see cref="BoxControlVM"/> objects.</summary>
    [CLSCompliant(true)]
    public interface IButtonGroupVM: IControlVM { }

    [CLSCompliant(true)]
    public interface ISplitButtonVM: IControlVM {
        IMenuVM   MenuVM   { get; }
    }

     /// <summary>The total interface exposed by <see cref="SplitToggleButtonVM"/> objects.</summary>
   [CLSCompliant(true)]
    public interface ISplitToggleButtonVM: ISplitButtonVM {
        IToggleVM ToggleVM { get; }
    }

     /// <summary>The total interface exposed by <see cref="SplitPressButtonVM"/> objects.</summary>
    [CLSCompliant(true)]
    public interface ISplitPressButtonVM: ISplitButtonVM {
        IButtonVM ButtonVM { get; }
    }

     /// <summary>The total interface exposed by <see cref="MenuVM"/> objects.</summary>
    [CLSCompliant(true)]
    public interface IMenuVM: IControlVM, IDescriptionableVM { }

    [CLSCompliant(true)]
    public interface IMenuSeparatorVM: IControlVM {
        string Title { get; }
    }
}
