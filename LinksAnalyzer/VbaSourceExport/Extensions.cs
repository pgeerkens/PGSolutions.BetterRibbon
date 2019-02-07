////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;

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

        /// <summary>.</summary>
        /// <param name="this"></param>
        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope")]
        internal static AccessWrapper NewAccessWrapper(this IApplication @this) {
            try {
                @this.DisplayAlerts = false;
                return new AccessWrapper();
            } finally {
                @this.DisplayAlerts = true;
            }
        }
    }
}
