////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using Microsoft.Office.Core;

namespace PGSolutions.RibbonDispatcher.ViewModels {
    /// <summary>The ViewModel for ButtonVM objects.</summary>
    internal class ButtonVM: AbstractControlVM<IButtonSource,IButtonVM>, IButtonVM,
            IActivatable<IButtonSource,IButtonVM>, ISizeableVM, IClickableVM, IImageableVM {
        public ButtonVM(string itemId) : base(itemId) { }

        /// <inheritdoc/>
        public virtual string Description => (Strings as IControlStrings2)?.Description ?? $"{Id} Description";

        #region IActivatable implementation
        /// <inheritdoc/>
        public override IButtonVM Attach(IButtonSource source) => Attach<ButtonVM>(source);

        /// <inheritdoc/>
        public override void Detach() { Clicked = null; base.Detach(); }
        #endregion

        #region IClickable implementation
        /// <inheritdoc/>
        public event ClickedEventHandler Clicked;

        /// <inheritdoc/>
        public virtual void OnClicked(IRibbonControl control) => Clicked?.Invoke(control);
        #endregion

        #region ISizeable implementation
        /// <inheritdoc/>
        public bool IsLarge => Source?.IsLarge ?? false;
        #endregion

        #region IImageable implementation
        /// <inheritdoc/>
        public ImageObject Image => Source?.Image ?? "MacroSecurity";

        /// <inheritdoc/>
        public bool ShowImage => Source?.ShowImage ?? (Source?.Image != null);

        /// <inheritdoc/>
        public bool ShowLabel => Source?.ShowLabel ?? true;
        #endregion
    }
}
