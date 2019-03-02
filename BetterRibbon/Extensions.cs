////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017-8 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

namespace PGSolutions.BetterRibbon {
    using PGSolutions.RibbonDispatcher.ViewModels;

    /// <summary>Extension methods for Excel objects.</summary>
    [CLSCompliant(true)]
    public static partial class Extensions {
        /// <inheritdoc/>
        public static ImageObject ToggleImage(this bool isPressed)
        => isPressed ? "TagMarkComplete" : "MarginsShowHide";
    }
}
