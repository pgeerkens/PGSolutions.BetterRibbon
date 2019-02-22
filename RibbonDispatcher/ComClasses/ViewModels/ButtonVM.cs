////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses.ViewModels {
    /// <summary>The ViewModel for ButtonVM objects.</summary>
    public class ButtonVM : AbstractControlVM<IRibbonButtonSource>, IRibbonButton,
            IActivatable<IRibbonButtonSource,ButtonVM>, ISizeable, IClickable, IImageable {
        internal ButtonVM(string itemId) : base(itemId) { }

        #region IActivatable implementation
        public new ButtonVM Attach(IRibbonButtonSource source) => Attach<ButtonVM>(source);

        public override void Detach() {
            Clicked = null;
            base.Detach();
        }
        #endregion

        #region IClickable implementation
        /// <summary>The Clicked event source for COM clients</summary>
        public event EventHandler Clicked;

        /// <summary>The callback from the Ribbon Dispatcher to initiate Clicked events on this control.</summary>
        public virtual void OnClicked(object sender, EventArgs e)
        => Clicked?.Invoke(this, EventArgs.Empty);
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
