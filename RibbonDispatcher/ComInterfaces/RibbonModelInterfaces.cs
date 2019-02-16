////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary></summary>
    [Description("")]    
    [ComVisible(true)]
    [CLSCompliant(false)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IRibbonGroupModel)]
    public interface IRibbonGroupModel {
        /// <summary>Gets the <see cref="IRibbonControlStrings"/> for this control.</summary>
        IRibbonControlStrings Strings {
            [Description("Gets the IRibbonControlStrings for this control.")]
            get;
        }

        /// <summary>Gets or sets whether the control is enabled.</summary>
        bool IsEnabled {
            [Description("Gets or sets whether the control is enabled.")]
            get; set;
        }
        /// <summary>Gets or sets whether the control is visible.</summary>
        bool IsVisible {
            [Description("Gets or sets whether the control is visible.")]
            get; set;
        }

        /// <summary>Gets whether or not inactive controls should be visible on the Ribbon.</summary>
        [Description("Gets whether or not inactive controls should be visible on the Ribbon.")]
        bool ShowInactive { get; }

        /// <summary>Sets whether or not inactive controls should be visible on the Ribbon.</summary>
        [Description("Sets whether or not inactive controls should be visible on the Ribbon.")]
        void SetShowInactive(bool showInactive);

        /// <summary>Attaches this control-model to the specified ribbon-control as data source and event sink.</summary>
        [Description("Attaches this control-model to the specified ribbon-control as data source and event sink.")]
        IRibbonGroupModel Attach(string controlId);

        /// <summary>Queues a request for this control to be refreshed.</summary>
        [Description("Queues a request for this control to be refreshed.")]
        void Invalidate();

        /// <summary>Detaches this Ribbon Group, and all child models, from their view-models.</summary>
        [Description("Detaches this Ribbon Group, and all child models, from their view-models.")]
        void Detach();
    }

    public interface IRibbonCommonSource {
        /// <summary>Gets the <see cref="IRibbonControlStrings"/> for this control.</summary>
        IRibbonControlStrings Strings { get; }

        /// <summary>Gets whether the control is enabled.</summary>
        bool IsEnabled { get; }

        /// <summary>Gets whether the control is visible.</summary>
        bool IsVisible { get; }

        bool ShowInactive { get; }

        void SetShowInactive(bool showInactive);
    }

    public interface IRibbonButtonSource : IRibbonCommonSource {
        /// <summary>Gets the <see cref="IRibbonControlStrings"/> for this control.</summary>
        new IRibbonControlStrings Strings { get; }

        /// <summary>Gets whether the control is enabled.</summary>
        new bool IsEnabled { get; }

        /// <summary>Gets whether the control is visible.</summary>
        new bool IsVisible { get; }

        new bool ShowInactive { get; }

        object Image       { get; }

        /// <summary>Gets whether the image for this control should be displayed when its size is {rdRegular}.</summary>
        bool ShowImage     { get; }

        /// <summary>Gets whether the label for this control should be displayed when its size is {rdRegular}.</summary>
        bool ShowLabel     { get; }

        bool IsLarge       { get; }
    }

    public interface IRibbonToggleSource : IRibbonButtonSource {
        /// <summary>Gets the <see cref="IRibbonControlStrings"/> for this control.</summary>
        new IRibbonControlStrings Strings { get; }

        /// <summary>Gets whether the control is enabled.</summary>
        new bool IsEnabled { get; }

        /// <summary>Gets whether the control is visible.</summary>
        new bool IsVisible { get; }

        new bool ShowInactive { get; }

        new object Image   { get; }

        new bool ShowImage { get; }

        new bool ShowLabel { get; }

        new bool IsLarge   { get; }

        bool IsPressed     { get; }
    }

    public interface IRibbonDropDownSource : IRibbonCommonSource, IEnumerable<ISelectableItem> {
        /// <summary>Gets the <see cref="IRibbonControlStrings"/> for this control.</summary>
        new IRibbonControlStrings Strings { get; }

        /// <summary>Gets whether the control is enabled.</summary>
        new bool IsEnabled { get; }

        /// <summary>Gets whether the control is visible.</summary>
        new bool IsVisible { get; }

        new bool ShowInactive { get; }

        int SelectedIndex { get; }

        ISelectableItem this[int index] { get; }

        int Count { get; }

        new IEnumerator<ISelectableItem> GetEnumerator();
    }

    public interface ISelectableItemSource:IRibbonCommonSource {
        /// <summary>Gets the <see cref="IRibbonControlStrings"/> for this control.</summary>
        new IRibbonControlStrings Strings { get; }

        /// <summary>Gets whether the control is enabled.</summary>
        new bool IsEnabled { get; }

        /// <summary>Gets whether the control is visible.</summary>
        new bool IsVisible { get; }

        new bool ShowInactive { get; }

        object Image { get; }

        /// <summary>Gets whether the image for this control should be displayed when its size is {rdRegular}.</summary>
        bool ShowImage { get; }

        /// <summary>Gets whether the label for this control should be displayed when its size is {rdRegular}.</summary>
        bool ShowLabel { get; }

        bool IsLarge { get; }
    }
}
