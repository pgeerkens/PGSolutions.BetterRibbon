////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

namespace PGSolutions.RibbonDispatcher.ControlMixins {
    /// <summary>TODO</summary>
    [CLSCompliant(true)]
    public delegate void ClickedEventHandler();

    internal static class ClickableMixin {
        public static void Clicked(this IClickableMixin mixin) => mixin.OnClicked();
    }
}
