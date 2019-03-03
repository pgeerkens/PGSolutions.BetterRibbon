////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////

namespace PGSolutions.RibbonDispatcher.ViewModels {
    public class MenuVM: AbstractContainerVM<IMenuSource,IMenuVM>, IMenuVM,
            IActivatable<IMenuSource,IMenuVM>, IImageableVM {
        internal MenuVM(ViewModelFactory factory, string itemId) : base(itemId) { }

        #region IActivatable implementation
        /// <summary>Attaches this control-model to the specified ribbon-control as data source and event sink.</summary>
        public override IMenuVM Attach(IMenuSource source) => Attach<MenuVM>(source);
        #endregion

        #region IImageable implementation
        /// <inheritdoc/>
        public IImageObject Image => Source?.Image ?? "MacroSecurity".ToImageObject();

        /// <inheritdoc/>
        public bool ShowImage => Source?.ShowImage ?? (Source?.Image != null);

        /// <inheritdoc/>
        public bool ShowLabel => Source?.ShowLabel ?? true;
        #endregion

        #region IDescriptionable implementation
        /// <inheritdoc/>
        public virtual string Description => (Strings as IControlStrings2)?.Description ?? $"{ControlId} Description";
        #endregion
    }
}
