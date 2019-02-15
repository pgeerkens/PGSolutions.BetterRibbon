using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using stdole;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    /// <summary>TODO</summary>
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable",
       Justification = "Public, Non-Creatable, class with exported Events.")]
    [Serializable]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ISelectableItem))]
    [Guid(Guids.SelectableItem)]
    public class SelectableItem : RibbonCommon<IRibbonCommonSource>, ISelectableItem, IImageable {
        /// <summary>TODO</summary>
        internal SelectableItem(string ItemId, IRibbonControlStrings strings, ImageObject Image) 
        : base(ItemId) {
            _strings = strings;
            _image = Image;
        }

        protected override IRibbonControlStrings Strings => _strings;
        private IRibbonControlStrings _strings;

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
