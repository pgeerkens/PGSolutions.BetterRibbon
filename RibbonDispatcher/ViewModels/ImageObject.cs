////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using stdole;

namespace PGSolutions.RibbonDispatcher.ViewModels {
    /// <summary></summary>
    [SuppressMessage("Microsoft.Performance", "CA1815:OverrideEqualsAndOperatorEqualsOnValueTypes",
            Justification ="Unnecessaty.")]
    [CLSCompliant(true)]
    public class ImageObject: IImageObject {
        // TODO - only used in BrandingModel constructor - necessary?
        public ImageObject(string imageMso)    => _image = imageMso;
        public ImageObject(IPictureDisp image) => _image = image;

        public   bool         IsMso     => ImageMso != null;
        public   string       ImageMso  => _image as string;
        public   IPictureDisp ImageDisp => _image as IPictureDisp;

        private  object _image { get; }

        [SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates",
                Justification = "Unnecessary - the existing properties achieve that.")]
        public static implicit operator ImageObject(string s) => new ImageObject(s);
    }

    public static partial class Extensions {
        public static object Image(this IImageObject @this)
        =>  @this.IsMso ? @this.ImageMso as object : @this.ImageDisp;

        public static IImageObject ToImageObject(this string @this) => new ImageObject(@this);
    }
}
