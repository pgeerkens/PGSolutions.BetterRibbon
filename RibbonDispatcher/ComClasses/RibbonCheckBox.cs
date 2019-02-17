////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.ComInterfaces;

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
    public class RibbonCheckBox : RibbonCommon<IRibbonToggleSource>, IRibbonToggle,
        IActivatable<RibbonCheckBox, IRibbonToggleSource>, IToggleable {
        internal RibbonCheckBox(string itemId) : base(itemId) { }

        #region IActivatable implementation
        RibbonCheckBox IActivatable<RibbonCheckBox, IRibbonToggleSource>.Attach(IRibbonToggleSource source)
        => Attach<RibbonCheckBox>(source);

        public override void Detach() {
            Toggled = null;
            base.Detach();
        }
        #endregion

        #region IToggleable implementation
        /// <summary>TODO</summary>
        public event ToggledEventHandler Toggled;

        /// <inheritdoc/>>
        public bool IsPressed => Source?.IsPressed ?? false;

        /// <inheritdoc/>>
        public virtual void OnToggled(object sender, bool isPressed) => Toggled?.Invoke(this,isPressed);
        #endregion

        #region ISizeable implementation
        /// <inheritdoc/>>
        public virtual bool IsSizeable => false;

        /// <summary>Gets or sets the preferred {RibbonControlSize} for the control.</summary>
        public virtual bool IsLarge => false;
        #endregion

        #region IImageable implementation
        /// <inheritdoc/>
        public virtual bool IsImageable => false;
        /// <inheritdoc/>
        public virtual ImageObject Image => null;

        /// <summary>Gets or sets whether the image for this control should be displayed when its size is {rdRegular}.</summary>
        public virtual bool ShowImage => false;

        /// <summary>Gets or sets whether the label for this control should be displayed when its size is {rdRegular}.</summary>
        public virtual bool ShowLabel => true;
        #endregion

        public override void Invalidate() => base.Invalidate();
    }
}
