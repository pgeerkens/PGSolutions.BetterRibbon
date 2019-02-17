////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

namespace PGSolutions.RibbonUtilities {
    internal static class NativeMethods {
        public static void KeyDown(this byte Vk) => Vk.keybd_event(0x10, KEYEVENTF_KEYDOWN, 0);
        public static void KeyUp(this byte Vk)   => Vk.keybd_event(0x10, KEYEVENTF_KEYUP, 0);

        private const uint KEYEVENTF_KEYDOWN = 0x0;
        private const uint KEYEVENTF_KEYUP   = 0x2;

        [SuppressMessage("Microsoft.Portability", "CA1901:PInvokeDeclarationsShouldBePortable", MessageId = "3")]
        [DllImport("user32.dll")]
        #pragma warning disable IDE1006 // Naming Styles - Matches name in external DLL
        private static extern void keybd_event(this byte bVk, byte bScan, uint dwFlags, int dwExtraInfo);
        #pragma warning restore IDE1006 // Naming Styles
    }
}
