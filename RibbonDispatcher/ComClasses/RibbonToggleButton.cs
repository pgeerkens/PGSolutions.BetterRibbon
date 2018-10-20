////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using stdole;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using Microsoft.Office.Core;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    /// <summary>The ViewModel for Ribbon ToggleButton objects.</summary>
    [Description("The ViewModel for Ribbon ToggleButton objects")]
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable",
       Justification = "Public, Non-Creatable, class with exported Events.")]
    [Serializable]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IToggledEvents))]
    [ComDefaultInterface(typeof(IRibbonToggle))]
    [Guid(Guids.RibbonToggleButton)]
    public class RibbonToggleButton : RibbonCheckBox, IRibbonToggle, ISizeable, IImageable {
        internal RibbonToggleButton(string itemId, IRibbonControlStrings strings, bool visible, bool enabled,
            RibbonControlSize size, ImageObject image, bool showImage, bool showLabel
        ) : base(itemId, strings, visible, enabled) {
            _size      = size;
            _image     = image;
            _showImage = showImage;
            _showLabel = showLabel;
        }

        #region IActivatable implementation
        public override void Detach() {
            SetImageMso("MacroSecurity");
            base.Detach();;
        }
        #endregion

        /// <inheritdoc/>>
        public override string Label => IsPressed || string.IsNullOrEmpty(AlternateLabel)
                                     ? base.Label ?? Id
                                     : AlternateLabel;

        #region ISizeable implementation
        /// <inheritdoc/>>
        public override bool IsSizeable => true;
        /// <summary>Gets or sets the preferred {RibbonControlSize} for the control.</summary>
        public override RibbonControlSize Size {
            get => _size;
            set { _size = value; Invalidate(); }
        }
        private RibbonControlSize _size;
        #endregion

        #region IImageable implementation
        /// <inheritdoc/>
        public override bool IsImageable => true;
        /// <inheritdoc/>
        public override object Image => _image.Image;
        private ImageObject _image;

        /// <summary>Gets or sets whether the image for this control should be displayed when its size is {rdRegular}.</summary>
        public override bool ShowImage {
            get => _showImage;
            set { _showImage = value; Invalidate(); }
        }
        private bool _showImage;

        /// <inheritdoc/>
        public override bool ShowLabel {
            get => _showLabel;
            set { _showLabel = value; Invalidate(); }
        }
        private bool _showLabel;

        /// <summary>Sets the displayable image for this control to the provided {IPictureDisp}</summary>
        public override void SetImageDisp(IPictureDisp Image) => _image = new ImageObject(Image);

        /// <summary>Sets the displayable image for this control to the named ImageMso image</summary>
        public override void SetImageMso(string ImageMso)     => _image = new ImageObject(ImageMso);
        #endregion
    }
}
