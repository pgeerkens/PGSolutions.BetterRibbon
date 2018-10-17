using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace PGSolutions.RibbonDispatcher.ComInterfaces {
    /// <summary>TODO</summary>
    [ComVisible(true)]
    [CLSCompliant(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid(Guids.IResourceManager)]
    public interface IResourceManager {
        /// <summary>Returns the {IRibbonTextLanguageControl} for the given {ControlId}.</summary>
        [DispId(DispIds.GetControlStrings)]
        [Description("Returns the IRibbonTextLanguageControl for the given ControlId.")]
        IRibbonControlStrings GetControlStrings(string ControlId);

        /// <summary>Returns the image(as an ImageMso string or an IPictureDisp) associated with the supplied name.</summary>
        [DispId(DispIds.GetImage)]
        [Description("Returns the image(as an ImageMso string or an IPictureDisp) associated with the supplied name.")]
        object GetImage(string Name);
    }

    internal static partial class DispIds {
        public const int GetControlStrings = 1;
        public const int GetImage          = 1 + GetControlStrings;
    }
}
