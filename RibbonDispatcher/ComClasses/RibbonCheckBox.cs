////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using stdole;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.Utilities;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    /// <summary>The ViewModel for Ribbon CheckBox objects.</summary>
    [Description("The ViewModel for Ribbon CheckBox objects.")]
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable",
        Justification = "Public, Non-Creatable, class with exported Events.")]
    [Serializable]
    [CLSCompliant(false)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IToggledEvents))]
    [ComDefaultInterface(typeof(IRibbonToggle))]
    [Guid(Guids.RibbonCheckBox)]
    public class RibbonCheckBox : RibbonCommon, IRibbonToggle, IActivatableControl<IRibbonCommon, bool>, IToggleable {
        internal RibbonCheckBox(string itemId, IRibbonControlStrings strings, bool visible, bool enabled
        ) : base(itemId, strings, visible, enabled) { }

        #region IActivatable implementation
        public IRibbonToggle Attach(Func<bool> getter) {
            base.Attach();
            Getter = getter;
            return this;
        }

        public override void Detach() {
            Toggled = null;
            Getter = () => false;
            base.Detach();
        }

        IRibbonCommon IActivatableControl<IRibbonCommon, bool>.Attach(Func<bool> getter) =>
            Attach(getter) as IRibbonCommon;
        void IActivatableControl<IRibbonCommon, bool>.Detach() => Detach();
        #endregion

        #region IToggleable implementation
        /// <summary>TODO</summary>
        public event ToggledEventHandler Toggled;

        /// <inheritdoc/>>
        public bool IsPressed => Getter?.Invoke() ?? false;

        /// <inheritdoc/>>
        public virtual void OnToggled(object sender, bool isPressed) => Toggled?.Invoke(this,isPressed);

        /// <summary>TODO</summary>
        private Func<bool> Getter { get; set; }
        #endregion

        #region ISizeable implementation
        /// <inheritdoc/>>
        public virtual bool IsSizeable => false;

        /// <summary>Gets or sets the preferred {RibbonControlSize} for the control.</summary>
        public virtual bool IsLarge {
            get => false;
            set { /* NO-OP */ }
        }
        #endregion

        #region IImageable implementation
        /// <inheritdoc/>
        public virtual bool IsImageable => false;
        /// <inheritdoc/>
        public virtual object Image => null;

        /// <summary>Gets or sets whether the image for this control should be displayed when its size is {rdRegular}.</summary>
        public virtual bool ShowImage {
            get => false;
            set { /* NO-OP */ }
        }

        /// <summary>Gets or sets whether the label for this control should be displayed when its size is {rdRegular}.</summary>
        public virtual bool ShowLabel {
            get => true;
            set { /* NO-OP */ }
        }

        /// <summary>Sets the displayable image for this control to the provided {IPictureDisp}</summary>
        public virtual void SetImageDisp(IPictureDisp Image) { /* NO-OP */ }

        /// <summary>Sets the displayable image for this control to the named ImageMso image</summary>
        public virtual void SetImageMso(string ImageMso)     { /* NO-OP */ }
        #endregion
    }
}
