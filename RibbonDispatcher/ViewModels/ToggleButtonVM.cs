////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////

namespace PGSolutions.RibbonDispatcher.ViewModels {
    /// <summary>The ViewModel for Ribbon ToggleButton objects.</summary>
    internal class ToggleButtonVM : CheckBoxVM, IToggleVM,
        IActivatable<IToggleSource,IToggleVM>, ISizeableVM, IImageableVM {
        internal ToggleButtonVM(string itemId) : base(itemId) { }

        #region IActivatable implementation
        /// <inheritdoc/>
        public override IToggleVM Attach(IToggleSource source) => Attach<ToggleButtonVM>(source);
        #endregion

        #region ISizeable implementation
        /// <inheritdoc/>
        public override bool IsLarge => Source?.IsLarge ?? false;
        #endregion

        #region IImageable implementation
        /// <inheritdoc/>
        public override IImageObject Image => Source?.Image ?? "MacroSecurity".ToImageObject();

        /// <inheritdoc/>
        public override bool ShowImage => Source?.ShowImage ?? (Source?.Image != null);

        /// <inheritdoc/>
        public override bool ShowLabel => Source?.ShowLabel ?? true;
        #endregion
    }
}
