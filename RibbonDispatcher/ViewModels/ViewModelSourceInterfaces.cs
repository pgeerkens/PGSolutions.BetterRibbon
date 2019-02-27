////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Diagnostics.CodeAnalysis;
using System.Collections.Generic;
using Microsoft.Office.Core;

namespace PGSolutions.RibbonDispatcher.ViewModels {
    using IStrings = IControlStrings;

    public delegate void ClickedEventHandler(IRibbonControl control);

    public delegate void ToggledEventHandler(IRibbonControl control, bool isPressed);

    public delegate void SelectionMadeEventHandler(IRibbonControl control, string selectedId, int selectedIndex);

    public delegate void EditedEventHandler(IRibbonControl control, string text);

    public interface ICanInvalidate {
        void Invalidate();
    }

    public interface IControlSource {
        /// <summary>Gets the <see cref="IRibbonControlStrings"/> for this control.</summary>
        IStrings Strings      { get; }

        /// <summary>Gets whether the control is enabled.</summary>
        bool     IsEnabled    { get; }

        /// <summary>Gets whether the control is visible.</summary>
        bool     IsVisible    { get; }

        /// <summary>.</summary>
        void     SetShowInactive(bool showInactive);
    }

    public interface ISizeSource {
        /// <summary>.</summary>
        bool        IsLarge   { get; }
    }

    public interface IImageSource {
        /// <summary>.</summary>
        ImageObject Image     { get; }

        /// <summary>Gets whether the image for this control should be displayed when its size is {rdRegular}.</summary>
        bool        ShowImage { get; }

        /// <summary>Gets whether the label for this control should be displayed when its size is {rdRegular}.</summary>
        bool        ShowLabel { get; }
    }

    public interface IToggleDataSource {
        /// <summary>.</summary>
        bool        IsPressed { get; }
    }

    public interface IEditDataSource {
        /// <summary>.</summary>
        string      Text      { get; }
    }

    public interface ISelectableSource {
        /// <summary>.</summary>
        int     SelectedIndex { get; }
        /// <summary>.</summary>
        string  SelectedId    { get; }
    }

    [SuppressMessage("Microsoft.Naming","CA1710:IdentifiersShouldHaveCorrectSuffix")]
    public interface IListDataSource : IEnumerable<ISelectableItemSource> {
        /// <summary>.</summary>
        int         Count     { get; }

        /// <summary>.</summary>
        ISelectableItemSource this[int index] { get; }

        /// <summary>.</summary>
        new IEnumerator<ISelectableItemSource> GetEnumerator();
    }


    public interface IEditBoxSource: IControlSource, IEditDataSource { }

    public interface IButtonSource : IControlSource, IImageSource, ISizeSource { }

    public interface IToggleSource : IControlSource, IToggleDataSource, IImageSource, ISizeSource { }

    [SuppressMessage("Microsoft.Naming", "CA1710:IdentifiersShouldHaveCorrectSuffix")]
    public interface IDropDownSource: IControlSource, ISelectableSource, IListDataSource { }

    public interface IStaticDropDownSource : IControlSource, ISelectableSource, IListDataSource { }

    [SuppressMessage("Microsoft.Naming", "CA1710:IdentifiersShouldHaveCorrectSuffix")]
    public interface IComboBoxSource : IControlSource, IEditBoxSource, IListDataSource { }

    public interface IStaticComboBoxSource : IControlSource, IEditBoxSource, IListDataSource { }

    public interface ISelectableItemSource: IControlSource, IImageSource, ISizeSource {
        string Id        { get; }
    }

    public interface ILabelSource: IControlSource, ISizeSource { }

    public interface IMenuSource: IControlSource, IImageSource { }
}
