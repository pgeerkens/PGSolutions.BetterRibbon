////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using Microsoft.Office.Core;
using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses.ViewModels {
    /// <summary>The ViewModel for ButtonVM objects.</summary>
    internal class ButtonVM : AbstractControlVM<IButtonSource>, IButtonVM,
            IActivatable<IButtonSource,IButtonVM>, ISizeableVM, IClickableVM, IImageableVM {
        internal ButtonVM(string itemId) : base(itemId) { }

        #region IActivatable implementation
        public new IButtonVM Attach(IButtonSource source) => Attach<ButtonVM>(source);

        public override void Detach() {
            Clicked = null;
            base.Detach();
        }
        #endregion

        #region IClickable implementation
        /// <summary>The Clicked event source for COM clients</summary>
        public event ClickedEventHandler Clicked;

        /// <summary>The callback from the Ribbon ModelFactory to initiate Clicked events on this control.</summary>
        public virtual void OnClicked(IRibbonControl control)
        => Clicked?.Invoke(control);
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
