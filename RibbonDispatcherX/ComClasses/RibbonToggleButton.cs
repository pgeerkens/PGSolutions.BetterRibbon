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
    /// <summary>The ViewModel for Ribbon ToggleButton objects.</summary>
    [Description("The ViewModel for Ribbon ToggleButton objects")]
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable",
       Justification = "Public, Non-Creatable, class with exported Events.")]
    [Serializable]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IToggledEvents))]
    [ComDefaultInterface(typeof(IRibbonToggleButton))]
    [Guid(Guids.RibbonToggleButton)]
    public class RibbonToggleButton : RibbonCommon, IRibbonToggleButton, IActivatableControl<IRibbonCommon, bool>,
        ISizeableMixin, IToggleableMixin, IImageableMixin {
        internal RibbonToggleButton(string itemId, IRibbonControlStrings strings, bool visible, bool enabled,
            RdControlSize size, ImageObject image, bool showImage, bool showLabel
        ) : base(itemId, strings, visible, enabled) {
            this.SetSize(size);
            this.SetImage(image);
            this.SetShowImage(showImage);
            this.SetShowLabel(showLabel);
        }

        #region IToggleable implementation
        private bool _isAttached    = false;

        public override bool IsEnabled => base.IsEnabled && _isAttached;
        public override bool IsVisible => base.IsVisible || ShowWhenInactive;

        public bool ShowWhenInactive { get; set; } = true;

        public IRibbonToggleButton Attach(Func<bool> getter) {
            _isAttached = true;
            this.SetGetter(getter);
            return this;
        }

        public void Detach() {
            _isAttached = false;
            this.SetGetter(() => false);
            SetLanguageStrings(RibbonTextLanguageControl.Empty);
            SetImageMso("MacroSecurity");
        }

        IRibbonCommon IActivatableControl<IRibbonCommon, bool>.Attach(Func<bool> getter) =>
            Attach(getter) as IRibbonCommon;
        void IActivatableControl<IRibbonCommon, bool>.Detach() => Detach();
        #endregion

        #region Publish IToggleableMixin to class default interface
        /// <summary>TODO</summary>
        public event ToggledEventHandler Toggled;

        /// <summary>TODO</summary>
        public virtual bool IsPressed   => this.GetPressed();

        /// <summary>TODO</summary>
        public override string Label    => this.GetLabel();

        /// <summary>TODO</summary>
        public virtual void OnToggled(bool IsPressed) => Toggled?.Invoke(IsPressed);

        /// <summary>TODO</summary>
        IRibbonControlStrings IToggleableMixin.LanguageStrings => Strings;
        #endregion

        #region Publish ISizeableMixin to class default interface
        /// <summary>Gets or sets the preferred {RdControlSize} for the control.</summary>
        public RdControlSize Size {
            get => this.GetSize();
            set => this.SetSize(value);
        }
        #endregion

        #region Publish IImageableMixin to class default interface
        /// <inheritdoc/>
        public object Image => this.GetImage();

        /// <summary>Gets or sets whether the image for this control should be displayed when its size is {rdRegular}.</summary>
        public bool ShowImage {
            get => this.GetShowImage();
            set => this.SetShowImage(value);
        }

        /// <summary>Gets or sets whether the label for this control should be displayed when its size is {rdRegular}.</summary>
        public bool ShowLabel {
            get => this.GetShowLabel();
            set => this.SetShowLabel(value);
        }

        /// <summary>Sets the displayable image for this control to the provided {IPictureDisp}</summary>
        public void SetImageDisp(IPictureDisp Image) => this.SetImage(Image);

        /// <summary>Sets the displayable image for this control to the named ImageMso image</summary>
        public void SetImageMso(string ImageMso)     => this.SetImage(ImageMso);
        #endregion
    }
}
