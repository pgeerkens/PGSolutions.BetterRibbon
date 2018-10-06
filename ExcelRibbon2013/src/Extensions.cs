////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

using PGSolutions.RibbonDispatcher.AbstractCOM;
using PGSolutions.RibbonDispatcher.ControlMixins;
using System.Collections.Generic;

namespace PGSolutions.ExcelRibbon2013 {
    internal static class Extensions {
        public static ClickedEventHandler DefaultButtonAction(this IRibbonButton sender) => sender.MsgBoxShow;

        [SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters", MessageId = "System.Windows.Forms.MessageBox.Show(System.String,System.String,System.Windows.Forms.MessageBoxButtons,System.Windows.Forms.MessageBoxIcon,System.Windows.Forms.MessageBoxDefaultButton,System.Windows.Forms.MessageBoxOptions)")]
        public static void MsgBoxShow(this IRibbonButton control) =>
            MessageBox.Show($"{control?.Id ?? "Unknown Button"} pressed.", RibbonViewModel.MsgBoxTitle,
                    MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, 0);

        /// <summary>Returns the text for the resource named <paramref name="resourceName"/>; else null if not found.</summary>
        public static string GetResourceText(this string resourceName) {
            var asm = Assembly.GetExecutingAssembly();
            using (var reader = ( from r in asm.GetManifestResourceNames()
                                  where string.Compare(resourceName, r, StringComparison.OrdinalIgnoreCase) == 0
                                  select new StreamReader(asm.GetManifestResourceStream(r))
                                ).FirstOrDefault() ) { return reader?.ReadToEnd(); }
        }

        public static void SetView(this IList<IRibbonButton> buttons, int selectedIndex) {
            foreach (var b in buttons) {
                b.ShowLabel = ((selectedIndex + 1) & 0x1) != 0x0;
                b.ShowImage = ((selectedIndex + 1) & 0x2) != 0x0;
            }
        }
    }
}
