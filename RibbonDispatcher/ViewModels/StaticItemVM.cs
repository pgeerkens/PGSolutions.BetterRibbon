////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////

namespace PGSolutions.RibbonDispatcher.ViewModels {
    /// <summary>TODO</summary>
    internal class StaticItemVM: AbstractControlVM<ISelectableItemSource>, ISelectableItemVM,
            IActivatable<ISelectableItemSource,ISelectableItemVM>, IImageableVM {
        /// <summary>TODO</summary>
        internal StaticItemVM(string ItemId, IControlStrings strings) : base(ItemId)
        => Strings = strings;

        protected override IControlStrings Strings { get; }

        #region IActivatable implementation
        /// <inheritdoc/>
        public new ISelectableItemVM Attach(ISelectableItemSource source) => Attach<StaticItemVM>(source);
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
