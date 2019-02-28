////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;
using System.Windows.Forms;

namespace PGSolutions.RibbonDispatcher {
    public static partial class Extensions {
        private const string Caption = "PGSolutions Ribbon Dispatcher";

        /// <summary>Returns the name of the calling method. </summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed",
                Justification = "Because that's just how it works!")]
        public static string CallerName([CallerMemberName] string memberName = "") => memberName;

        /// <summary>.</summary>
        /// <param name="message"></param>
        public static void MsgBoxShow(this string message) => message.MsgBoxShow(Caption);

        /// <summary>.</summary>
        /// <param name="message"></param>
        /// <param name="caption"></param>
        public static void MsgBoxShow(this string message, string caption) =>
            MessageBox.Show($"{message}.", caption, MessageBoxButtons.OK, MessageBoxIcon.Information);

        public  static string Format2(this Version version) =>
            $"{version?.Major}.{version?.Minor}.{version?.Build}.{version?.Revision}";
        public  static string Format(this Version version) => Format2(version) +
            $"({version?.Build.FormatVersionDate()} " +
            $"{version?.Revision.FormatVersionTime()} UTC)";
        private static string FormatVersionDate(this int dayNo) =>
            new DateTime(2000, 1, 1).AddDays(dayNo).ToUniversalTime().ToString("yyyy-MM-dd");
        private static string FormatVersionTime(this int halfSeconds) =>
            new DateTime(2000, 1, 1).AddSeconds(2 * halfSeconds).ToUniversalTime().ToString("HH:mm:ss");
    }
}
