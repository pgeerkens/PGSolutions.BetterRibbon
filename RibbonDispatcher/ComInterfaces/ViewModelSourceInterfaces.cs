////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Diagnostics.CodeAnalysis;
using System.Collections.Generic;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    using IStrings = IControlStrings;

    public interface IInvalidate {
        void Invalidate();
    }

    public interface IRibbonCommonSource {
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

    public interface IButtonSource : IRibbonCommonSource {
        ///// <summary>Gets the <see cref="IControlStrings"/> for this control.</summary>
        //new IStrings Strings  { get; }

        ///// <summary>Gets whether the control is enabled.</summary>
        //new bool IsEnabled    { get; }

        ///// <summary>Gets whether the control is visible.</summary>
        //new bool IsVisible    { get; }

        ///// <summary>.</summary>
        //new bool ShowInactive { get; }

        /// <summary>.</summary>
        ImageObject Image     { get; }

        /// <summary>Gets whether the image for this control should be displayed when its size is {rdRegular}.</summary>
        bool ShowImage        { get; }

        /// <summary>Gets whether the label for this control should be displayed when its size is {rdRegular}.</summary>
        bool ShowLabel        { get; }

        /// <summary>.</summary>
        bool IsLarge          { get; }
    }

    public interface IRibbonToggleSource : IButtonSource {
        ///// <summary>Gets the <see cref="IControlStrings"/> for this control.</summary>
        //new IStrings Strings  { get; }

        ///// <summary>Gets whether the control is enabled.</summary>
        //new bool IsEnabled    { get; }

        ///// <summary>Gets whether the control is visible.</summary>
        //new bool IsVisible    { get; }

        ///// <summary>.</summary>
        //new bool ShowInactive { get; }

        ///// <summary>.</summary>
        //new ImageObject Image { get; }

        ///// <summary>.</summary>
        //new bool ShowImage    { get; }

        ///// <summary>.</summary>
        //new bool ShowLabel    { get; }

        ///// <summary>.</summary>
        //new bool IsLarge      { get; }

        /// <summary>.</summary>
        bool IsPressed        { get; }
    }

    [SuppressMessage("Microsoft.Naming", "CA1710:IdentifiersShouldHaveCorrectSuffix")]
    public interface IDropDownSource : IRibbonCommonSource, IEnumerable<ISelectableItem> {
        ///// <summary>Gets the <see cref="IControlStrings"/> for this control.</summary>
        //new IStrings Strings  { get; }

        ///// <summary>Gets whether the control is enabled.</summary>
        //new bool IsEnabled    { get; }

        ///// <summary>Gets whether the control is visible.</summary>
        //new bool IsVisible    { get; }

        ///// <summary>.</summary>
        //new bool ShowInactive { get; }

        /// <summary>.</summary>
        int SelectedIndex     { get; }

        /// <summary>.</summary>
        int Count             { get; }

        /// <summary>.</summary>
        ISelectableItem this[int index] { get; }

        /// <summary>.</summary>
        new IEnumerator<ISelectableItem> GetEnumerator();
    }

    public interface ISelectableItemSource: IRibbonCommonSource {
        ///// <summary>Gets the <see cref="IControlStrings"/> for this control.</summary>
        //new IStrings Strings      { get; }

        ///// <summary>Gets whether the control is enabled.</summary>
        //new bool     IsEnabled    { get; }

        ///// <summary>Gets whether the control is visible.</summary>
        //new bool     IsVisible    { get; }

        ///// <summary>.</summary>
        //new bool     ShowInactive { get; }

        /// <summary>.</summary>
        ImageObject  Image        { get; }

        /// <summary>Gets whether the image for this control should be displayed when its size is {rdRegular}.</summary>
        bool         ShowImage    { get; }

        /// <summary>Gets whether the label for this control should be displayed when its size is {rdRegular}.</summary>
        bool         ShowLabel    { get; }

        /// <summary>.</summary>
        bool         IsLarge      { get; }
    }

    public interface IEditBoxSource: IRibbonCommonSource {
        ///// <summary>Gets the <see cref="IControlStrings"/> for this control.</summary>
        //new IStrings Strings      { get; }

        ///// <summary>Gets whether the control is enabled.</summary>
        //new bool     IsEnabled    { get; }

        ///// <summary>Gets whether the control is visible.</summary>
        //new bool     IsVisible    { get; }

        ///// <summary>.</summary>
        //new bool     ShowInactive { get; }

        /// <summary>.</summary>
        string       Text         { get; }
    }

    [SuppressMessage("Microsoft.Naming", "CA1710:IdentifiersShouldHaveCorrectSuffix")]
    public interface IComboBoxSource: IDropDownSource, IEditBoxSource {
        ///// <summary>Gets the <see cref="IControlStrings"/> for this control.</summary>
        //new IStrings Strings       { get; }

        ///// <summary>Gets whether the control is enabled.</summary>
        //new bool     IsEnabled     { get; }

        ///// <summary>Gets whether the control is visible.</summary>
        //new bool     IsVisible     { get; }

        ///// <summary>.</summary>
        //new bool     ShowInactive  { get; }

        ///// <summary>.</summary>
        //new int      SelectedIndex { get; }

        ///// <summary>.</summary>
        //new int      Count         { get; }

        ///// <summary>.</summary>
        //new string   Text          { get; }

        ///// <summary>.</summary>
        //new ISelectableItem this[int index] { get; }

        ///// <summary>.</summary>
        //new IEnumerator<ISelectableItem> GetEnumerator();
    }
}
