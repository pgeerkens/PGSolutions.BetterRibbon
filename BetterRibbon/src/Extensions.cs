////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Windows.Forms;

namespace PGSolutions.RibbonUtilities.LinksAnalyzer {
    internal static partial class Extensions {
        public static void ShowMsgString(this string message, string caption = "",
                MessageBoxIcon icon = MessageBoxIcon.None) =>
            MessageBox.Show(message, caption,
                MessageBoxButtons.OK,
                icon,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);
    }
}
