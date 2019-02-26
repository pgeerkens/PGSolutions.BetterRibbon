////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using stdole;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary></summary>
    [Description("")]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IToggleModel)]
    public interface IToggleModel {
        /// <summary>Gets the <see cref="IRibbonControlStrings"/> for this control.</summary>
        IControlStrings Strings {
            [Description("Gets the IControlStrings for this control.")]
            get;
        }

        /// <summary>Gets or sets whether the control is enabled.</summary>
        [Description(".")]
        bool   IsEnabled {
            [Description("Gets or sets whether the control is enabled.")]
            get; set; }
        /// <summary>Gets or sets whether the control is visible.</summary>
        [Description(".")]
        bool   IsVisible {
            [Description("Gets or sets whether the control is visible.")]
            get; set; }
        /// <summary>.</summary>
        [Description(".")]
        bool   IsLarge   { get; set; }
        /// <summary>.</summary>
        [Description(".")]
        ImageObject Image     { get; set; }
        /// <summary>.</summary>
        [Description(".")]
        bool   ShowImage { get; set; }
        /// <summary>.</summary>
        [Description(".")]
        bool   ShowLabel { get; set; }

        /// <summary>.</summary>
        [Description(".")]
        bool   IsPressed { get; set; }

        /// <summary>Attaches this control-model to the specified ribbon-control as data source and event sink.</summary>
        [Description("Attaches this control-model to the specified ribbon-control as data source and event sink.")]
        IToggleModel Attach(string controlId);

        /// <summary>Queues a request for this control to be refreshed.</summary>
        [Description("Queues a request for this control to be refreshed.")]
        void Invalidate();

        /// <summary>.</summary>
        [Description(".")]
        void SetImageDisp(IPictureDisp image);
        /// <summary>.</summary>
        [Description(".")]
        void SetImageMso(string imageMso);
    }
}
