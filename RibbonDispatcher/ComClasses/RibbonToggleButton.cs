////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
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
            bool isLarge, ImageObject image, bool showImage, bool showLabel
        ) : base(itemId, strings, visible, enabled) {
            _size      = isLarge.ControlSize();
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

        #region IToggleable implementation
        /// <inheritdoc/>>
        public override string Label => IsPressed || string.IsNullOrEmpty(AlternateLabel)
                                     ? base.Label ?? Id
                                     : AlternateLabel;
        #endregion

        #region ISizeable implementation
        /// <inheritdoc/>>
        public override bool IsSizeable => true;
        /// <summary>Gets or sets the preferred {RibbonControlSize} for the control.</summary>
        public override bool IsLarge {
            get => _size == RibbonControlSize.RibbonControlSizeLarge;
            set { _size = value.ControlSize() ; Invalidate(); }
        }
        private RibbonControlSize _size;
        #endregion

        #region IImageable implementation
        /// <inheritdoc/>
        public override bool IsImageable => true;
        /// <inheritdoc/>
        public override object Image => _image.Image;

        private void SetImage(ImageObject image) {
            _image = image;
            Invalidate();
        }
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
        public void SetImageDisp(IPictureDisp Image) => SetImage(new ImageObject(Image));

        /// <summary>Sets the displayable image for this control to the named ImageMso image</summary>
        public void SetImageMso(string ImageMso) => SetImage(new ImageObject(ImageMso));
        #endregion
    }
}
