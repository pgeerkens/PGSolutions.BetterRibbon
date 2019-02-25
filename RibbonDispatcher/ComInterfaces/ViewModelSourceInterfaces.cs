////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Diagnostics.CodeAnalysis;
using System.Collections.Generic;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    using IStrings = IControlStrings;

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
        bool     ShowInactive { get; }

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

    public interface ISelectableSource: IControlSource, IEnumerable<ISelectableItemModel> {
        /// <summary>.</summary>
        int         Count     { get; }

        /// <summary>.</summary>
        ISelectableItemModel this[int index] { get; }

        /// <summary>.</summary>
        new IEnumerator<ISelectableItemModel> GetEnumerator();
    }

    [SuppressMessage("Microsoft.Naming", "CA1710:IdentifiersShouldHaveCorrectSuffix")]
    public interface IDropDownSource : ISelectableSource {
        /// <summary>.</summary>
        int     SelectedIndex { get; }
    }

    public interface IEditBoxSource: IControlSource {
        /// <summary>.</summary>
        string      Text      { get; }
    }

    [SuppressMessage("Microsoft.Naming", "CA1710:IdentifiersShouldHaveCorrectSuffix")]
    public interface IComboBoxSource: ISelectableSource, IEditBoxSource {
    }

    public interface ISelectableItemSource: IControlSource {
        /// <summary>.</summary>
        ImageObject Image { get; }

        /// <summary>Gets whether the image for this control should be displayed when its size is {rdRegular}.</summary>
        bool ShowImage { get; }

        /// <summary>Gets whether the label for this control should be displayed when its size is {rdRegular}.</summary>
        bool ShowLabel { get; }

        /// <summary>.</summary>
        bool IsLarge { get; }
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
