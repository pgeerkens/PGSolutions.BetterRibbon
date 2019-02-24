////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Runtime.InteropServices;
using stdole;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    /// <summary>TODO</summary>
    /// <remarks>
    /// </remarks>
    [Serializable]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IResourceLoader))]
    [Guid(Guids.ResourceLoader)]
    public class ResourceLoader : IResourceLoader, IResourceManager {
        /// <summary>Creates a new empty ControlStrings collection.</summary>
        internal ResourceLoader() {
            _controls = new Dictionary<string, IControlStrings>();
            _images   = new Dictionary<string, IPictureDisp>();
        }

        Dictionary<string, IControlStrings>  _controls;
        Dictionary<string, IPictureDisp>           _images;

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
            _controls.Add(ControlId, new ControlStrings(
                    Label           ?? ControlId,
                    ScreenTip       ?? $"{ControlId} ScreenTip",
                    SuperTip        ?? $"{ControlId} SuperTip",
                    KeyTip          ?? "",
                    AlternateLabel  ?? $"{ControlId} Alternate",
                    Description     ?? $"{ControlId} Description"));
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
        public IControlStrings GetControlStrings(string ControlId) =>
            _controls.FirstOrDefault(i => i.Key == ControlId).Value;
        /// <inheritdoc/>
        public object GetImage(string Name) =>
            _images.FirstOrDefault(i => i.Key == Name).Value;

        /// <inheritdoc/>
        public IControlStrings this[string ControlId] => _controls.FirstOrDefault(i => i.Key == ControlId).Value;
    }
}
