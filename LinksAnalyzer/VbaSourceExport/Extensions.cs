////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

namespace PGSolutions.RibbonUtilities.VbaSourceExport {
    public static partial class Extensions {
        internal static void InvokeWithShiftKey(this Action action) {
            const byte VK_LSHIFT = 0xA0;  // left shift key
            try {
                VK_LSHIFT.KeyDown();
                action();
            } finally {
                VK_LSHIFT.KeyUp();
            }
        }
    }
}
