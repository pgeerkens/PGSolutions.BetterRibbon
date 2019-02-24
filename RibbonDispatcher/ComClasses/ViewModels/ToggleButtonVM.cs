////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses.ViewModels {
    /// <summary>The ViewModel for Ribbon ToggleButton objects.</summary>
    internal class ToggleButtonVM : CheckBoxVM, IToggleControlVM,
        IActivatable<IToggleSource,ToggleButtonVM>, ISizeableVM, IImageableVM {
        internal ToggleButtonVM(string itemId) : base(itemId) { }

        #region IActivatable implementation
        /// <inheritdoc/>
        public new ToggleButtonVM Attach(IToggleSource source) => Attach<ToggleButtonVM>(source);
        #endregion

        #region IToggleable implementation
        /// <inheritdoc/>>
        public override string Label => !IsPressed || string.IsNullOrEmpty(Strings?.AlternateLabel)
                                     ? base.Label
                                     : AlternateLabel;
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
