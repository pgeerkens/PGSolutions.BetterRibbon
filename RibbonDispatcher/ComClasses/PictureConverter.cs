////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.Windows.Forms;
using stdole;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    [SuppressMessage("Microsoft.Design", "CA1052:StaticHolderTypesShouldBeSealed")]
    [SuppressMessage("Microsoft.Performance", "CA1812:AvoidUninstantiatedInternalClasses",
            Justification = "False positive - static methods ARE accessed.")]
    public class PictureConverter:AxHost {
        private PictureConverter() : base(string.Empty) { }

        public static IPictureDisp ImageToPictureDisp(Image image)
        => GetIPictureDispFromPicture(image) as IPictureDisp;

        public static IPictureDisp IconToPictureDisp(Icon icon)
        => ImageToPictureDisp(icon?.ToBitmap());
    }
}
