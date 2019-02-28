////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////

namespace PGSolutions.RibbonDispatcher.ViewModels {
    /// <summary>TODO</summary>
    public class StaticItemVM: AbstractControlVM<ISelectableItemSource>, IStaticItemVM,
            IActivatable<ISelectableItemSource,IStaticItemVM>, IImageableVM {
        /// <summary>TODO</summary>
        internal StaticItemVM(string ItemId, IControlStrings strings) : base(ItemId)
        => Strings = strings;

        protected override IControlStrings Strings { get; }

        #region IActivatable implementation
        /// <inheritdoc/>
        public new IStaticItemVM Attach(ISelectableItemSource source) => Attach<StaticItemVM>(source);
        #endregion

        #region IImageableVM implementation
        /// <inheritdoc/>
        public ImageObject Image => Source?.Image ?? "MacroSecurity";

        /// <inheritdoc/>
        public bool ShowImage => Source?.ShowImage ?? (Source?.Image != null);

        /// <inheritdoc/>
        public bool ShowLabel => Source?.ShowLabel ?? true;
        #endregion
    }
}
