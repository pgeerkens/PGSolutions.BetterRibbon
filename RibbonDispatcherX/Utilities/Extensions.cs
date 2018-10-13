////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ControlMixins;
using System.Collections.Generic;

namespace PGSolutions.RibbonDispatcher.Utilities {
    public static class Extensions {
        /// <summary>Displays a {MessageBox} identifying the (supplied) source {IRibbonButton}/</summary>
        public static ClickedEventHandler DefaultButtonAction(this IRibbonButton sender) => sender.MsgBoxShow;

        public static void MsgBoxShow<TControl>(this TControl control) where TControl : IRibbonButton => MsgBoxShow(control, null);
        [SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "System.Windows.Forms.MessageBox.Show(System.String,System.String,System.Windows.Forms.MessageBoxButtons,System.Windows.Forms.MessageBoxIcon,System.Windows.Forms.MessageBoxDefaultButton,System.Windows.Forms.MessageBoxOptions)")]

        public static void MsgBoxShow<TControl>(this TControl control, string details) where TControl : IRibbonButton =>
            MessageBox.Show($"{control?.Id ?? "Unknown Button"} pressed {details??""}.", "PGSolutions Ribbon Dispatcher",
                    MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);

        /// <summary>Returns the text for the resource named <paramref name="resourceName"/>; else null if not found.</summary>
        public static string GetResourceText(this string resourceName) {
            var asm = Assembly.GetExecutingAssembly();
            using (var reader = ( from r in asm.GetManifestResourceNames()
                                  where string.Compare(resourceName, r, StringComparison.OrdinalIgnoreCase) == 0
                                  select new StreamReader(asm.GetManifestResourceStream(r))
                                ).FirstOrDefault() ) { return reader?.ReadToEnd(); }
        }

        [Flags]
        public enum LabelImageOptions {
            None      = 0x0,
            ShowLabel = 0x1,
            ShowImage = 0x2,
            ShowBoth  = ShowLabel | ShowImage
        }
        public static int IndexFromLabelImageDisplay(this LabelImageOptions value) => (int)(value - 1);
        public static LabelImageOptions IndexToLabelImageDisplay(this int value) => (LabelImageOptions)(value + 1);

        /// <summary>Set the display of all supplied {IRibbonImageable}s as per the supplied {displayFlags}.</summary>
        public static void SetDisplay<T>(this IList<T> buttons, int index) where T : IRibbonImageable
            => buttons.SetDisplay(index.IndexToLabelImageDisplay());

        /// <summary>Set the display of all supplied {IRibbonImageable}s as per the supplied {displayFlags}.</summary>
        public static void SetDisplay<T>(this IList<T> buttons, LabelImageOptions displayOptions) where T: IRibbonImageable {
            foreach (var b in buttons) {
                b.ShowLabel = displayOptions.HasFlag(LabelImageOptions.ShowLabel);
                b.ShowImage = displayOptions.HasFlag(LabelImageOptions.ShowImage);
            }
        }
    }
}
