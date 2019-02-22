﻿////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    [CLSCompliant(false)]
    public interface IRibbonImageable {
        /// <summary>Returns the current Image for the control as either a {string} naming an MsoImage or an {IPictureDisp}.</summary>
        [Description("Returns the current Image for the control as either a {string} naming an MsoImage or an {IPictureDisp}.")]
        ImageObject Image       { get; }
        /// <summary>Gets or sets whether to show the control's image; ignored by Large controls.</summary>
        [Description("Gets or sets whether to show the control's image; ignored by Large controls.")]
        bool ShowImage          { get; }
        /// <summary>Gets or sets whether to show the control's label; ignored by Large controls.</summary>
        [Description("Gets or sets whether to show the control's label; ignored by Large controls.")]
        bool ShowLabel          { get; }
    }
}
