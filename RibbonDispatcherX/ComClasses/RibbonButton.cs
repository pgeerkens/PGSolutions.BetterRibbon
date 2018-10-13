////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using stdole;

using PGSolutions.RibbonDispatcher.ControlMixins;
using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.Utilities;

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
    public class RibbonButton : RibbonCommon, IRibbonButton, IActivatableControl<IRibbonCommon>,
        ISizeableMixin, IClickableMixin, IImageableMixin {
        internal RibbonButton(string itemId, IResourceManager mgr, bool visible, bool enabled, RdControlSize size,
                ImageObject image, bool showImage, bool showLabel) : base(itemId, mgr, visible, enabled) {
            this.SetSize(size);
            this.SetImage(image);
            this.SetShowImage(showImage);
            this.SetShowLabel(showLabel);
            _preferredSize = size;
        }

        #region IActivatable implementation
        private bool _isAttached    = false;
        private bool _enableVisible = true;
        private readonly RdControlSize _preferredSize;

        public override bool IsEnabled => base.IsEnabled && _isAttached;
        public override bool IsVisible => base.IsVisible && _enableVisible;

        public IRibbonButton Attach() {
            this.SetSize(_preferredSize);
            _isAttached = true;
            _enableVisible = true;
            return this;
        }

        public void Detach() => Detach(true);
        public void Detach(bool enableVisible) {
            _enableVisible = enableVisible;
            _isAttached = false;
            SetLanguageStrings(RibbonTextLanguageControl.Empty);
            SetImageMso("MacroSecurity");
            this.SetSize(RdControlSize.rdRegular);
        }

        IRibbonCommon IActivatableControl<IRibbonCommon>.Attach() => Attach() as IRibbonCommon;
        void IActivatableControl<IRibbonCommon>.Detach() => Detach();
        #endregion

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
        public virtual void OnClicked() => Clicked?.Invoke();
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
