////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using stdole;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    public struct ImageObject {
        public ImageObject(string imageMso)    => _image = imageMso;
        public ImageObject(IPictureDisp image) => _image = image;

        public object Image => IsMso ? ImageMso as object : ImageDisp;
        public bool         IsMso => ImageMso != null;
        public string       ImageMso  => _image as string;
        public IPictureDisp ImageDisp => _image as IPictureDisp;

        private object _image { get; }

        public static implicit operator ImageObject(string s) => new ImageObject(s);
    }
}
