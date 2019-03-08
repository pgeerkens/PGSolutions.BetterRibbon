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
        /// <summary>Returns a <see cref="IControlStrings"/> for the given {ControlId}.</summary>
        [DispId(1), Description("Returns a IControlStrings for the given ControlId.")]
        IControlStrings GetControlStrings(string ControlId);

        /// <summary>Returns a <see cref="IControlStrings2"/> for the given {ControlId}.</summary>
        /// <remarks>
        /// The ribbon controls that support a Description attribute, and the associated 
        /// GetDescription callback, require an <see cref="IControlStrings2"/> initialization. These
        /// include:
        ///     Button, ToggleButton, CheckBox, Menu, Gallery, and DynamicMenu.
        /// </remarks>
        [DispId(2), Description(
@"Returns a IControlStrings2 for the given ControlId.

The ribbon controls that support a Description attribute, and the associated 
GetDescription callback, require an IControlStrings2 initialization. These
include:
    Button, ToggleButton, CheckBox, Menu, Gallery, and DynamicMenu."
        )]
        IControlStrings2 GetControlStrings2(string ControlId);

        /// <summary>Returns the image(as an ImageMso string or an IPictureDisp) associated with the supplied name.</summary>
        [DispId(3), Description("Returns the image(as an ImageMso string or an IPictureDisp) associated with the supplied name.")]
        IPictureDisp GetImage(string Name);
    }
}
