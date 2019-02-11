////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using stdole;

using Microsoft.Office.Core;
using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    public static partial class Extensions {
        internal static void SetSizeablel(this ISizeable target, ISizeable source)
        => target.IsLarge = source.IsLarge;

        internal static void SetImageable(this IImageable target, IImageable source) {
            target.ShowImage = source.ShowImage;
            target.ShowLabel = source.ShowLabel;
            if (source.Image is string) target.SetImageMso(source.Image as string);
            if (source.Image is IPictureDisp) target.SetImageDisp(source.Image as IPictureDisp);
        }

        public static RibbonControlSize ControlSize(this bool isLarge)
            => isLarge ? RibbonControlSize.RibbonControlSizeLarge
                       : RibbonControlSize.RibbonControlSizeRegular;
    }
}
