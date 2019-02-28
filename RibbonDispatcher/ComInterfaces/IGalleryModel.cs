////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using stdole;

using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary></summary>
    [Description("")]
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IGalleryModel)]
    public interface IGalleryModel {
        /// <summary>Gets the <see cref="IRibbonControlStrings"/> for this control.</summary>
        IControlStrings Strings {
            [Description("Gets the IControlStrings for this control.")]
            get;
        }

        /// <summary>Gets or sets whether the control is enabled.</summary>
        bool IsEnabled {
            [Description("Gets or sets whether the control is enabled.")]
            get; set; }

        /// <summary>Gets or sets whether the control is visible.</summary>
        bool IsVisible {
            [Description("Gets or sets whether the control is visible.")]
            get; set; }

        /// <summary>Gets or sets the (zero-based) integer of the selected item.</summary>
        int SelectedIndex {
            [Description("Gets or sets the (zero-based) integer of the selected item.")]
            get; set; }

        /// <summary>Adds the specified <see cref="ISelectableItem"/> to the available options in the drop-down list.</summary>
        IGalleryModel AddSelectableModel(IStaticItemVM selectableModel);

        /// <summary>Gets or sets the height in pixels for items.</summary>
        int ItemHeight {
            [Description("Gets or sets the height in pixels for items.")]
            get; set; }

        /// <summary>Gets or sets the width in pixels for items.</summary>
        int ItemWidth{
            [Description("Gets or sets the width in pixels for items.")]
            get; set; }

        #region IImageable implementation
        /// <summary>.</summary>
        ImageObject Image {
        [Description(".")]
            get; set; }

        bool ShowImage {
            [Description("Gets or sets the width in pixels for items.")]
            get; set; }

        bool ShowLabel {
            [Description("Gets or sets the width in pixels for items.")]
            get; set; }

        /// <summary>.</summary>
        [Description(".")]
        void SetImageDisp(IPictureDisp image);
        /// <summary>.</summary>
        [Description(".")]
        void SetImageMso(string imageMso);
        #endregion

        /// <summary>Attaches this control-model to the specified ribbon-control as data source and event sink.</summary>
        [Description("Attaches this control-model to the specified ribbon-control as data source and event sink.")]
        IGalleryModel Attach(string controlId);

        /// <summary>.</summary>
        [Description(".")]
        void Detach();

        /// <summary>Queues a request for this control to be refreshed.</summary>
        [Description("Queues a request for this control to be refreshed.")]
        void Invalidate();
    }
}
