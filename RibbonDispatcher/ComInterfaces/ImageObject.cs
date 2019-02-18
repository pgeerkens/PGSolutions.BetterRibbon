////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using stdole;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    [SuppressMessage("Microsoft.Performance", "CA1815:OverrideEqualsAndOperatorEqualsOnValueTypes")]
    /// <summary></summary>
    [Description("")]
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable",
            Justification = "Public, Non-Creatable, class with exported Events.")]
    [SuppressMessage("Microsoft.Performance", "CA1815:OverrideEqualsAndOperatorEqualsOnValueTypes",
            Justification ="Unnecessaty.")]
    [ComVisible(true)]
    [CLSCompliant(false)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IImageObject))]
    [Guid(Guids.ImageObject)]
    public class ImageObject:IImageObject {
        public ImageObject(string imageMso)    => _image = imageMso;
        public ImageObject(IPictureDisp image) => _image = image;

        public object       Image     => IsMso ? ImageMso as object : ImageDisp;
        public bool         IsMso     => ImageMso != null;
        public string       ImageMso  => _image as string;
        public IPictureDisp ImageDisp => _image as IPictureDisp;

        private object _image { get; }

        [SuppressMessage("Microsoft.Usage", "CA2225:OperatorOverloadsHaveNamedAlternates",
                Justification = "Unnecessary - the existing properties achieve that.")]
        public static implicit operator ImageObject(string s) => new ImageObject(s);
    }
}
