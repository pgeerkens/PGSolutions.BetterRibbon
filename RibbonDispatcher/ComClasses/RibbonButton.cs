////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using stdole;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.Utilities;
using Microsoft.Office.Core;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    /// <summary>The ViewModel for RibbonButton objects.</summary>
    [Description("The ViewModel for Ribbon Button objects.")]
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable",
        Justification = "Public, Non-Creatable, class with exported Events.")]
    [Serializable]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IClickedEvents))]
    [ComDefaultInterface(typeof(IRibbonButton))]
    [Guid(Guids.RibbonButton)]
    public class RibbonButton : RibbonCommon, IRibbonButton, IActivatableControl<IRibbonButton>,
        ISizeable, IClickable, IImageable {
        internal RibbonButton(string itemId, IRibbonControlStrings strings, bool visible, bool enabled,
                RibbonControlSize size, ImageObject image, bool showImage, bool showLabel
        ) : base(itemId, strings, visible, enabled) {
            _size      = size;
            _image     = image;
            _showImage = showImage;
            _showLabel = showLabel;
        }

        #region IActivatable implementation
        public override void Detach() {
            Clicked = null;
            SetImageMso("MacroSecurity");
            base.Detach();
        }

        IRibbonButton IActivatableControl<IRibbonButton>.Attach() => Attach() as IRibbonButton;
        void IActivatableControl<IRibbonButton>.Detach() => Detach();
        #endregion

        #region IClickable implementation
        /// <summary>The Clicked event source for COM clients</summary>
        public event ClickedEventHandler Clicked;

        /// <summary>The callback from the Ribbon Dispatcher to initiate Clicked events on this control.</summary>
        public virtual void OnClicked(object sender) => Clicked?.Invoke(this);
        #endregion

        #region ISizeable implementation
        /// <inheritdoc/>
        public RibbonControlSize Size {
            get => _size;
            set { _size = value; Invalidate(); }
        }
        private RibbonControlSize _size;
        #endregion

        #region IImageable implementation
        /// <inheritdoc/>
        public bool IsImageable => true;
        /// <inheritdoc/>
        public object Image => _image.Image;
        private ImageObject _image;

        /// <summary>Gets or sets whether the image for this control should be displayed when its size is {rdRegular}.</summary>
        public bool ShowImage {
            get => _showImage;
            set { _showImage = value; Invalidate(); }
        }
        private bool _showImage;

        /// <inheritdoc/>
        public bool ShowLabel {
            get => _showLabel;
            set { _showLabel = value; Invalidate(); }
        }
        private bool _showLabel;

        /// <summary>Sets the displayable image for this control to the provided {IPictureDisp}</summary>
        public void SetImageDisp(IPictureDisp Image) => _image = new ImageObject(Image);

        /// <summary>Sets the displayable image for this control to the named ImageMso image</summary>
        public void SetImageMso(string ImageMso)     => _image = new ImageObject(ImageMso);
        #endregion
    }
}
