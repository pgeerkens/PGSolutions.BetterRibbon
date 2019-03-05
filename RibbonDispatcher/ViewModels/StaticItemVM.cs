////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////

namespace PGSolutions.RibbonDispatcher.ViewModels {
    /// <summary>TODO</summary>
    public class StaticItemVM: AbstractControlVM<ISelectableItemSource,IStaticItemVM>, IStaticItemVM,
            IActivatable<ISelectableItemSource,IStaticItemVM>, IImageableVM {
        /// <summary>TODO</summary>
        internal StaticItemVM(string ItemId, IControlStrings strings) : base(ItemId) {
            Label     = strings.Label;
            ScreenTip = strings.ScreenTip;
            SuperTip  = strings.SuperTip;
            KeyTip    = strings.KeyTip;
        }

        /// <inheritdoc/>
        public override string Label     { get; }
        /// <inheritdoc/>
        public override string ScreenTip { get; }
        /// <inheritdoc/>
        public override string SuperTip  { get; }
        /// <inheritdoc/>
        public override string KeyTip    { get; }

        #region IActivatable implementation
        /// <inheritdoc/>
        public override IStaticItemVM Attach(ISelectableItemSource source) => Attach<StaticItemVM>(source);
        #endregion

        #region IImageableVM implementation
        /// <inheritdoc/>
        public IImageObject Image => Source?.Image ?? "MacroSecurity".ToImageObject();

        /// <inheritdoc/>
        public bool ShowImage => Source?.ShowImage ?? (Source?.Image != null);

        /// <inheritdoc/>
        public bool ShowLabel => Source?.ShowLabel ?? true;
        #endregion
    }
}
