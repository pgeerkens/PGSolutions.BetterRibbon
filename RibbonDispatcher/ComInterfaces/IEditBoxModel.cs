﻿////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary></summary>
    [Description("")]
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IEditBoxModel)]
    public interface IEditBoxModel {
        event EditedEventHandler Edited;

        /// <summary>Gets or sets the content of the <see cref="EditBoxVM"/>. Default value.</summary>
        [DispId(12)]
        string Text {
            [Description("Gets or sets the content of the EditBoxVM. Default value.")]
            get; set;
        }

        #region IActivable implementation
        /// <summary>Attaches this control-model to the specified ribbon-control as data source and event sink.</summary>
        [DispId(1),Description("Attaches this control-model to the specified ribbon-control as data source and event sink.")]
        IEditBoxModel Attach(string controlId);

        /// <summary>.</summary>
        [DispId(2),Description(".")]
        void Detach();

        /// <summary>Queues a request for this control to be refreshed.</summary>
        [DispId(3),Description("Queues a request for this control to be refreshed.")]
        void Invalidate();
        #endregion

        #region IControl implementation
        /// <summary>Gets the <see cref="IControlStrings"/> for this control.</summary>
        [DispId(4)]
        string Label {
            [Description("Gets the IControlStrings for this control.")]
            get; set;
        }
        /// <summary>Gets the ScreenTip (concise hover-help) for this control.</summary>
        [DispId(17)]
        string ScreenTip {
            [Description("Gets the ScreenTip (concise hover-help) for this control.")]
            get; set;
        }
        /// <summary>Gets the SuperTip (expanded hover-help) for this control.</summary>
        [DispId(18)]
        string SuperTip {
            [Description("Gets the SuperTip (expanded hover-help) for this control.")]
            get; set;
        }
        /// <summary>Gets the KeyTip (keyboard shortcut) for this control.</summary>
        [DispId(19)]
        string KeyTip {
            [Description("Gets the KeyTip (keyboard shortcut) for this control.")]
            get; set;
        }
        /// <summary>Gets or sets whether the control is enabled.</summary>
        [DispId(5)]
        bool IsEnabled {
            [Description("Gets or sets whether the control is enabled.")]
            get; set;
        }
        /// <summary>Gets or sets whether the control is visible.</summary>
        [DispId(6)]
        bool IsVisible {
            [Description("Gets or sets whether the control is visible.")]
            get; set;
        }
        #endregion

        #region IImageable implementation
        /// <summary>Returns ths current image for this control as either a <<see cref="string"/> or <see cref="IPictureDisp"/>.</summary>
        [DispId(8)]
        IImageObject Image {
            [Description("Returns ths current image for this control as either a string or IPictureDisp.")]
            get; }
        /// <summary>Gets or sets Whether this control displays an image.</summary>
        [DispId(9)]
        bool   ShowImage  {
            [Description("Gets or sets Whether this control displays an image.")]
            get; set; }
        /// <summary>Gets or sets whether this control displays a label.</summary>
        [DispId(10)]
        bool   ShowLabel  {
            [Description("Gets or sets whether this control displays a label.")]
            get; set; }

        /// <summary>Sets the image for this control to the MCO image as named.</summary>
        [DispId(11),Description("Sets the current image for this control to the provided IPictureDisp.")]
        IEditBoxModel SetImage(IImageObject image);
        #endregion
    }
}
