////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses.ViewModels {
    internal class MenuVM: AbstractContainerVM<IMenuSource>, IMenuVM,
            IActivatable<IMenuSource,IMenuVM>, IImageableVM {
        public MenuVM(ViewModelFactory factory, string itemId) : base(factory, itemId) { }

        #region IActivatable implementation
        /// <summary>Attaches this control-model to the specified ribbon-control as data source and event sink.</summary>
        public new IMenuVM Attach(IMenuSource source) => Attach<MenuVM>(source);
        #endregion

        #region IImageable implementation
        /// <inheritdoc/>
        public ImageObject Image => Source?.Image ?? "MacroSecurity";

        /// <inheritdoc/>
        public bool ShowImage => Source?.ShowImage ?? (Source?.Image != null);

        /// <inheritdoc/>
        public bool ShowLabel => Source?.ShowLabel ?? true;
 
        /// <inheritdoc/>
        public virtual string Description => (Strings as IControlStrings2)?.Description ?? $"{Id} Description";
        #endregion
    }
}
