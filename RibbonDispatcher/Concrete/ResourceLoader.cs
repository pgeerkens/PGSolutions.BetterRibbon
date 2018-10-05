using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Runtime.InteropServices;
using stdole;

using PGSolutions.RibbonDispatcher.AbstractCOM;

namespace PGSolutions.RibbonDispatcher.Concrete {
    /// <summary>TODO</summary>
    /// <remarks>
    /// </remarks>
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable",
       Justification = "Public, Non-Creatable, class with exported Events.")]
    [Serializable]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IResourceLoader))]
    [Guid(Guids.ResourceLoader)]
    public class ResourceLoader : IResourceLoader {
        /// <summary>Creates a new empty ControlStrings collection.</summary>
        internal ResourceLoader() {
            _controls = new Dictionary<string, IRibbonTextLanguageControl>();
            _images   = new Dictionary<string, IPictureDisp>();
        }

        Dictionary<string, IRibbonTextLanguageControl>  _controls;
        Dictionary<string, IPictureDisp>    _images;

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [SuppressMessage("Microsoft.Globalization", "CA1305:SpecifyIFormatProvider", MessageId = "System.String.Format(System.String,System.Object)")]
        public string AddControlStrings(string ControlId,
            string Label            = null,
            string ScreenTip        = null,
            string SuperTip         = null,
            string AlternateLabel   = null,
            string Description      = null,
            string KeyTip           = null
        )
        {
            _controls.Add(ControlId, new RibbonTextLanguageControl(
                    Label           ?? ControlId,
                    ScreenTip       ?? ControlId + " ScreenTip",
                    SuperTip        ?? ControlId + " SuperTip",
                    KeyTip          ?? "",
                    AlternateLabel  ?? ControlId + " Alternate",
                    Description     ?? ControlId + " Description"));
            return ControlId;
        }

        /// <inheritdoc/>
        [SuppressMessage("Microsoft.Design", "CA1026:DefaultParametersShouldNotBeUsed", Justification = "Matches COM usage.")]
        [SuppressMessage("Microsoft.Globalization", "CA1305:SpecifyIFormatProvider", MessageId = "System.String.Format(System.String,System.Object)")]
        public string AddImage(string ImageId, IPictureDisp image) {
            _images.AddNotNull(ImageId, image);
            return ImageId;
        }

        /// <inheritdoc/>
        public IRibbonTextLanguageControl this[string ControlId] => _controls.FirstOrDefault(i => i.Key == ControlId).Value;
    }
}
