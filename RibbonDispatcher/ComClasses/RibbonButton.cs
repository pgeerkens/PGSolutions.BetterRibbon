////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.ComInterfaces;

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
    public class RibbonButton : RibbonCommon<IRibbonButtonSource>, IRibbonButton,
            IActivatable<RibbonButton,IRibbonButtonSource>, ISizeable, IClickable, IImageable {
        internal RibbonButton(string itemId) : base(itemId) { }

        #region IActivatable implementation
        RibbonButton IActivatable<RibbonButton, IRibbonButtonSource>.Attach(IRibbonButtonSource source)
        => Attach<RibbonButton>(source);

        public override void Detach() {
            Clicked = null;
            base.Detach();
        }
        #endregion

        #region IClickable implementation
        /// <summary>The Clicked event source for COM clients</summary>
        public event ClickedEventHandler Clicked;

        /// <summary>The callback from the Ribbon Dispatcher to initiate Clicked events on this control.</summary>
        public virtual void OnClicked() => Clicked?.Invoke(this);

        /// <summary>The callback from the Ribbon Dispatcher to initiate Clicked events on this control.</summary>
        public virtual void OnClicked(object sender) => Clicked?.Invoke(this);
        #endregion

        #region ISizeable implementation
        /// <inheritdoc/>
        public bool IsLarge => Source?.IsLarge ?? true;
        #endregion

        #region IImageable implementation
        /// <inheritdoc/>
        public bool IsImageable => true;

        /// <inheritdoc/>
        public ImageObject Image => Source?.Image ?? "MacroSecurity";

        /// <inheritdoc/>
        public bool ShowImage => Source?.ShowImage ?? true;

        /// <inheritdoc/>
        public bool ShowLabel => Source?.ShowLabel ?? true;
        #endregion
    }
}
