////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary>TODO</summary>
    [CLSCompliant(true)]
    public interface IResourceManager {
        /// <summary>Returns the {IRibbonTextLanguageControl} for the given {ControlId}.</summary>
        [Description("Returns the IRibbonTextLanguageControl for the given ControlId.")]
        IControlStrings GetControlStrings(string ControlId);

        /// <summary>Returns the {IRibbonTextLanguageControl} for the given {ControlId}.</summary>
        [Description("Returns the IRibbonTextLanguageControl for the given ControlId.")]
        IControlStrings2 GetControlStrings2(string ControlId);

        /// <summary>Returns the image(as an ImageMso string or an IPictureDisp) associated with the supplied name.</summary>
        [Description("Returns the image(as an ImageMso string or an IPictureDisp) associated with the supplied name.")]
        object GetImage(string Name);
    }
}
