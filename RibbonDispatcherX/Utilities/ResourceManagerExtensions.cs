////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2017 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.Globalization;
using System.Resources;
using System.Windows.Forms;
using Microsoft.Office.Core;
using stdole;

namespace PGSolutions.RibbonDispatcher.Utilities {
    /// <summary>TODO</summary>
    public static class ResourceManagerExtensions {
        /// <summary>TODO</summary>
        public static string GetCurrentUIString(this ResourceManager resourceManager, string name)
            => resourceManager?.GetString(name, CultureInfo.CurrentUICulture) ?? "";

        /// <summary>TODO</summary>
        public static string GetInvariantString(this ResourceManager resourceManager, string name)
            => resourceManager?.GetString(name, CultureInfo.InvariantCulture) ?? "";

        /// <summary>TODO</summary>
        public static IPictureDisp GetResourceIcon(this ResourceManager resourceManager, string iconName) {
            using (var icon = resourceManager?.GetObject(iconName, CultureInfo.InvariantCulture) as Icon) {
                return icon == null ? null : PictureConverter.IconToPictureDisp(icon);
            }
        }

        /// <summary>TODO</summary>
        public static IPictureDisp GetResourceImage(this ResourceManager resourceManager, string imageName) {
            using (var image = resourceManager?.GetObject(imageName, CultureInfo.InvariantCulture) as Image) {
                return (image == null) ? null : PictureConverter.ImageToPictureDisp(image);
            }
        }

        /// <summary>Returns a string as the ID of the supplied control suffixed with ' Unknown'.</summary>
        public static string Unknown(this IRibbonControl Control) => Unknown(Control?.Id, "Unknown");

        /// <summary>Returns a string as the ID of the supplied control suffixed with ' Unknown'.</summary>
        public static string Unknown(this string controlId) => Unknown(controlId, "Unknown");

        /// <summary>Returns a string as the ID of the supplied control suffixed with the supplied string.</summary>
        public static string Unknown(this IRibbonControl Control, string suffix) => Unknown(Control?.Id, suffix);

        /// <summary>Returns a string as the ID of the supplied control suffixed with the supplied string.</summary>
        public static string Unknown(this string controlId, string suffix)
            => string.Format(CultureInfo.InvariantCulture, $"'{controlId ?? ""}' {suffix}");

        /// <summary>TODO</summary>
        public static IPictureDisp ImageToPictureDisp(this Image image)
            => PictureConverter.ImageToPictureDisp(image) as IPictureDisp;

        /// <summary>TODO</summary>
        public static IPictureDisp IconToPictureDisp(this Icon icon)
            => PictureConverter.IconToPictureDisp(icon);

        [SuppressMessage("Microsoft.Performance", "CA1812:AvoidUninstantiatedInternalClasses",
                Justification="False positive - static methods ARE accessed.")]
        internal class PictureConverter : AxHost {
            private PictureConverter() : base(string.Empty) { }

            public static IPictureDisp ImageToPictureDisp(Image image) => GetIPictureDispFromPicture(image) as IPictureDisp;

            public static IPictureDisp IconToPictureDisp(Icon icon) => ImageToPictureDisp(icon.ToBitmap());
        }
    }
}
