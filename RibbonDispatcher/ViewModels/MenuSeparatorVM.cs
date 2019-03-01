////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////

namespace PGSolutions.RibbonDispatcher.ViewModels {
    internal class MenuSeparatorVM: AbstractControlVM<IMenuSeparatorSource,IMenuSeparatorVM>, IMenuSeparatorVM,
             IActivatable<IMenuSeparatorSource,IMenuSeparatorVM> {
        public MenuSeparatorVM(string itemId) : base(itemId) { }

        #region IActivatable implementation
        /// <summary>Attaches this control-model to the specified ribbon-control as data source and event sink.</summary>
        public override IMenuSeparatorVM Attach(IMenuSeparatorSource source) => Attach<MenuSeparatorVM>(source);
        #endregion

        #region ISizeable implementation
        /// <inheritdoc/>
        public string Title => Source?.Title ?? "";
        #endregion

        protected override bool DefaultShowInactive { get => true; set { } }
    }
}
