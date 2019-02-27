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

    public interface IButtonSource : IControlSource {
        /// <summary>.</summary>
        ImageObject Image     { get; }

        /// <summary>Gets whether the image for this control should be displayed when its size is {rdRegular}.</summary>
        bool        ShowImage { get; }

        /// <summary>Gets whether the label for this control should be displayed when its size is {rdRegular}.</summary>
        bool        ShowLabel { get; }

        /// <summary>.</summary>
        bool        IsLarge   { get; }
    }

    public interface IToggleSource : IButtonSource {
        /// <summary>.</summary>
        bool        IsPressed { get; }
    }

    [SuppressMessage("Microsoft.Naming","CA1710:IdentifiersShouldHaveCorrectSuffix")]
    public interface ISelectableSource: IControlSource, IEnumerable<ISelectableItemSource> {
        /// <summary>.</summary>
        int         Count     { get; }

        /// <summary>.</summary>
        ISelectableItemSource this[int index] { get; }

        /// <summary>.</summary>
        new IEnumerator<ISelectableItemSource> GetEnumerator();
    }

    public interface IStaticListSource {
        /// <summary>.</summary>
        int     SelectedIndex { get; }
        /// <summary>.</summary>
        string  SelectedId    { get; }
    }

    [SuppressMessage("Microsoft.Naming", "CA1710:IdentifiersShouldHaveCorrectSuffix")]
    public interface IDropDownSource : ISelectableSource, IStaticListSource {
    }

    public interface IStaticDropDownSource : IControlSource, IStaticListSource {
    }

    public interface IEditBoxSource: IControlSource {
        /// <summary>.</summary>
        string      Text      { get; }
    }

    [SuppressMessage("Microsoft.Naming", "CA1710:IdentifiersShouldHaveCorrectSuffix")]
    public interface IComboBoxSource: ISelectableSource, IEditBoxSource { }

    public interface ISelectableItemSource: IControlSource {
        string Id        { get; }

        /// <summary>Gets whether the image for this control should be displayed when its size is {rdRegular}.</summary>
        bool   ShowImage { get; }

        /// <summary>Gets whether the label for this control should be displayed when its size is {rdRegular}.</summary>
        bool   ShowLabel { get; }

        /// <summary>.</summary>
        bool   IsLarge   { get; }

        /// <summary>.</summary>
        ImageObject Image { get; }
    }

    public interface ILabelSource: IControlSource {
        /// <summary>.</summary>
        bool IsLarge { get; }
    }

    public interface IMenuSource: IControlSource {
        /// <summary>.</summary>
        ImageObject Image { get; }

        /// <summary>Gets whether the image for this control should be displayed when its size is {rdRegular}.</summary>
        bool ShowImage { get; }

        /// <summary>Gets whether the label for this control should be displayed when its size is {rdRegular}.</summary>
        bool ShowLabel { get; }
    }
}
