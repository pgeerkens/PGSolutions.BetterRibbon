////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses.ViewModels {
    /// <summary>The ViewModel for Ribbon ToggleButton objects.</summary>
    public class ToggleButtonVM : CheckBoxVM, IToggleButtonVM,
        IActivatable<IRibbonToggleSource,ToggleButtonVM>, ISizeable, IImageable {
        internal ToggleButtonVM(string itemId) : base(itemId) { }

        #region IActivatable implementation
        public new ToggleButtonVM Attach(IRibbonToggleSource source)
        => Attach<ToggleButtonVM>(source);
        #endregion

        #region IToggleable implementation
        /// <inheritdoc/>>
        public override string Label => IsPressed || string.IsNullOrEmpty(AlternateLabel)
                                     ? base.Label ?? Id
                                     : AlternateLabel;
        #endregion

        #region ISizeable implementation
        /// <inheritdoc/>>
        public override bool IsSizeable => true;
        /// <inheritdoc/>
        public override bool IsLarge => Source?.IsLarge ?? true;
        #endregion

        #region IImageable implementation
        /// <inheritdoc/>
        public override bool IsImageable => true;

        /// <inheritdoc/>
        public override ImageObject Image => Source?.Image ?? "MacroSecurity";

        /// <inheritdoc/>
        public override bool ShowImage => Source?.ShowImage ?? true;

        /// <inheritdoc/>
        public override bool ShowLabel => Source?.ShowLabel ?? true;
        #endregion
    }
}
