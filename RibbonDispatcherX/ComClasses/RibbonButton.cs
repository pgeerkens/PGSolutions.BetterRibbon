using System;
using System.Runtime.InteropServices;
using System.Diagnostics.CodeAnalysis;
using stdole;

using PGSolutions.RibbonDispatcher.ControlMixins;
using PGSolutions.RibbonDispatcher.AbstractCOM;
using System.ComponentModel;

namespace PGSolutions.RibbonDispatcher.ConcreteCOM {
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
    public class RibbonButton : RibbonCommon, IRibbonButton,
        ISizeableMixin, IClickableMixin, IImageableMixin {
        internal RibbonButton(string itemId, IResourceManager mgr, bool visible, bool enabled, RdControlSize size,
                ImageObject image, bool showImage, bool showLabel) : base(itemId, mgr, visible, enabled) {
            this.SetSize(size);
            this.SetImage(image);
            this.SetShowImage(showImage);
            this.SetShowLabel(showLabel);
        }

        #region Publish ISizeableMixin to class default interface
        /// <inheritdoc/>
        public RdControlSize Size {
            get => this.GetSize();
            set => this.SetSize(value);
        }
        #endregion

        #region Publish IClickableMixin to class default interface
        /// <summary>The Clicked event source for COM clients</summary>
        public event ClickedEventHandler Clicked;

        /// <summary>The callback from the Ribbon Dispatcher to initiate Clicked events on this control.</summary>
        public void OnClicked() => Clicked?.Invoke();
        #endregion

        #region Publish IImageableMixin to class default interface
        /// <inheritdoc/>
        public object Image => this.GetImage();

        /// <inheritdoc/>
        public bool ShowImage {
            get => this.GetShowImage();
            set => this.SetShowImage(value);
        }

        /// <inheritdoc/>
        public bool ShowLabel {
            get => this.GetShowLabel();
            set => this.SetShowLabel(value);
        }

        /// <inheritdoc/>
        public void SetImageDisp(IPictureDisp Image) => this.SetImage(Image);

        /// <inheritdoc/>
        public void SetImageMso(string ImageMso)     => this.SetImage(ImageMso);
        #endregion
    }
}
