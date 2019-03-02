////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Diagnostics.CodeAnalysis;
using System.Collections.Generic;

namespace PGSolutions.RibbonDispatcher.ViewModels {
    using IStrings = IControlStrings;

    #region Component interfaces
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
        IImageObject Image     { get; }

        /// <summary>Gets whether the image for this control should be displayed when its size is {rdRegular}.</summary>
        bool         ShowImage { get; }

        /// <summary>Gets whether the label for this control should be displayed when its size is {rdRegular}.</summary>
        bool         ShowLabel { get; }
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
    public interface IListDataSource {
        /// <summary>.</summary>
        IReadOnlyList<IStaticItemVM> Items { get; }
    }

    public interface IGridSizeSource {
        int ItemHeight { get; }
        int ItemWidth  { get; }
    }
    #endregion

    #region Object Source interfaces
    public interface IImageSizeSource: IControlSource, IImageSource, ISizeSource { }

    public interface IButtonSource : IImageSizeSource { }

    public interface IToggleSource : IImageSizeSource, IToggleDataSource { }

    public interface ISelectableItemSource:  IControlSource, IStaticItemVM { }

    public interface IEditBoxSource: IControlSource, IEditDataSource { }

    [SuppressMessage("Microsoft.Naming", "CA1710:IdentifiersShouldHaveCorrectSuffix")]
    public interface IDropDownSource: IControlSource, ISelectableSource, IListDataSource { }

    [SuppressMessage("Microsoft.Naming","CA1710:IdentifiersShouldHaveCorrectSuffix")]
    public interface IStaticDropDownSource : IControlSource, ISelectableSource { }

    [SuppressMessage("Microsoft.Naming", "CA1710:IdentifiersShouldHaveCorrectSuffix")]
    public interface IComboBoxSource : IControlSource, IEditBoxSource, IListDataSource { }

    [SuppressMessage("Microsoft.Naming","CA1710:IdentifiersShouldHaveCorrectSuffix")]
    public interface IStaticComboBoxSource : IControlSource, IEditBoxSource { }

    [SuppressMessage("Microsoft.Naming","CA1710:IdentifiersShouldHaveCorrectSuffix")]
    public interface IGallerySource : IControlSource, IGridSizeSource, ISelectableSource, IListDataSource { }

    [SuppressMessage("Microsoft.Naming","CA1710:IdentifiersShouldHaveCorrectSuffix")]
    public interface IStaticGallerySource: IControlSource, IGridSizeSource, ISelectableSource { }

    public interface ILabelControlSource: IControlSource, ISizeSource { }

    public interface IMenuSource: IControlSource, IImageSource { }

    public interface IMenuSeparatorSource: IControlSource {
        string Title { get; }
    }
    #endregion
}
