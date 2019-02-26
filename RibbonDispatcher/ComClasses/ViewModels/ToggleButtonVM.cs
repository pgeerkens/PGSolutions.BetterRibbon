////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses.ViewModels {
    /// <summary>The ViewModel for Ribbon ToggleButton objects.</summary>
    internal class ToggleButtonVM : CheckBoxVM, IToggleVM,
        IActivatable<IToggleSource,ToggleButtonVM>, ISizeableVM, IImageableVM {
        internal ToggleButtonVM(string itemId) : base(itemId) { }

        #region IActivatable implementation
        /// <inheritdoc/>
        public new ToggleButtonVM Attach(IToggleSource source) => Attach<ToggleButtonVM>(source);
        #endregion

        #region ISizeable implementation
        /// <inheritdoc/>
        public override bool IsLarge => Source?.IsLarge ?? false;
        #endregion

        #region IImageable implementation
        /// <inheritdoc/>
        public override ImageObject Image => Source?.Image ?? "MacroSecurity";

        /// <inheritdoc/>
        public override bool ShowImage => Source?.ShowImage ?? (Source?.Image != null);

        /// <inheritdoc/>
        public override bool ShowLabel => Source?.ShowLabel ?? true;
        #endregion
    }
}
