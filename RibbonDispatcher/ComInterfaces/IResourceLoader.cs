////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using stdole;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary>The default COM interface exposed by {ResourceLoader} objects.</summary>
    [Description("The default COM interface exposed by {ResourceLoader} objects.")]
        [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IResourceLoader)]
    public interface IResourceLoader {
        /// <summary>Returns the {IRibbonTextLanguageControl} for the given {ControlId}.</summary>
        [DispId(1), Description("Returns the IRibbonTextLanguageControl for the given ControlId.")]
        IControlStrings GetControlStrings(string ControlId);

        /// <summary>Returns the {IRibbonTextLanguageControl} for the given {ControlId}.</summary>
        [DispId(2), Description("Returns the IRibbonTextLanguageControl for the given ControlId.")]
        IControlStrings2 GetControlStrings2(string ControlId);

        /// <summary>Returns the image(as an ImageMso string or an IPictureDisp) associated with the supplied name.</summary>
        [DispId(3), Description("Returns the image(as an ImageMso string or an IPictureDisp) associated with the supplied name.")]
        IPictureDisp GetImage(string Name);
    }
}
