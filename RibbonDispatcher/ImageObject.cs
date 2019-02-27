////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Diagnostics.CodeAnalysis;
using stdole;

namespace PGSolutions.RibbonDispatcher.ViewModels {
    /// <summary></summary>
    [SuppressMessage("Microsoft.Performance", "CA1815:OverrideEqualsAndOperatorEqualsOnValueTypes",
            Justification ="Unnecessaty.")]
    public class ImageObject {
        public ImageObject(string imageMso)    => _image = imageMso;
        public ImageObject(IPictureDisp image) => _image = image;

        public object       Image     => IsMso ? ImageMso as object : ImageDisp;
        public bool         IsMso     => ImageMso != null;
        public string       ImageMso  => _image as string;
        public IPictureDisp ImageDisp => _image as IPictureDisp;

        private object _image { get; }

        public static ImageObject Empty => new ImageObject(null as IPictureDisp);

        [SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates",
                Justification = "Unnecessary - the existing properties achieve that.")]
        public static implicit operator ImageObject(string s) => new ImageObject(s);
    }
}
