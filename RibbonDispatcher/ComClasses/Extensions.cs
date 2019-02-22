////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Forms;

using Microsoft.Office.Core;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    public static partial class Extensions {
        private const string Caption = "PGSolutions Ribbon Dispatcher";

        public static RibbonControlSize ControlSize(this bool isLarge)
            => isLarge ? RibbonControlSize.RibbonControlSizeLarge
                       : RibbonControlSize.RibbonControlSizeRegular;

        /// <summary>Returns the name of the calling method. </summary>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Because that's just how it works!")]
        public static string CallerName([CallerMemberName] string memberName = "") => memberName;

        /// <summary>Displays a {MessageBox} identifying the (supplied) source {IButtonVM}/</summary>
        [SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "System.Windows.Forms.MessageBox.Show(System.String,System.String,System.Windows.Forms.MessageBoxButtons,System.Windows.Forms.MessageBoxIcon,System.Windows.Forms.MessageBoxDefaultButton,System.Windows.Forms.MessageBoxOptions)")]
        public static void DefaultButtonAction(object sender) =>
            $"{(sender as IButtonVM)?.Id ?? "Unknown Button"} pressed.".MsgBoxShow();

        public static void MsgBoxShow(this string message) => message.MsgBoxShow(Caption);

        [SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "System.Windows.Forms.MessageBox.Show(System.String,System.String,System.Windows.Forms.MessageBoxButtons,System.Windows.Forms.MessageBoxIcon,System.Windows.Forms.MessageBoxDefaultButton,System.Windows.Forms.MessageBoxOptions)")]
        public static void MsgBoxShow(this string message, string caption) =>
            MessageBox.Show($"{message}.", caption, MessageBoxButtons.OK, MessageBoxIcon.Information);

        /// <summary>Returns the text for the resource named <paramref name="resourceName"/>; else null if not found.</summary>
        public static string GetResourceText(this string resourceName) {
            var asm = Assembly.GetExecutingAssembly();
            using (var reader = (from r in asm.GetManifestResourceNames()
                                 where string.Compare(resourceName, r, StringComparison.OrdinalIgnoreCase) == 0
                                 select new StreamReader(asm.GetManifestResourceStream(r))
                                ).FirstOrDefault()) { return reader?.ReadToEnd(); }
        }

        public static string Format2(this Version version) =>
            $"{version?.Major}.{version?.Minor}.{version?.Build}.{version?.Revision}";
        public static string Format(this Version version) => Format2(version) +
            $"({version?.Build.FormatVersionDate()} " +
            $"{version?.Revision.FormatVersionTime()} UTC)";
        private static string FormatVersionDate(this int dayNo) =>
            new DateTime(2000, 1, 1).AddDays(dayNo).ToUniversalTime().ToString("yyyy-MM-dd");
        private static string FormatVersionTime(this int halfSeconds) =>
            new DateTime(2000, 1, 1).AddSeconds(2 * halfSeconds).ToUniversalTime().ToString("HH:mm:ss");

        [Flags]
        public enum LabelImageOptions {
            None = 0x0,
            ShowLabel = 0x1,
            ShowImage = 0x2,
            ShowBoth = ShowLabel | ShowImage
        }
        public static int IndexFromLabelImageDisplay(this LabelImageOptions value) => (int)(value - 1);
        [SuppressMessage("Microsoft.Usage", "CA2233:OperationsShouldNotOverflow", MessageId = "value+1")]
        public static LabelImageOptions IndexToLabelImageDisplay(this int value) => (LabelImageOptions)(value + 1);
    }
}
