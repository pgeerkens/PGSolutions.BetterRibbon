////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;

namespace PGSolutions.RibbonDispatcher.ViewModels {
    [SuppressMessage("Microsoft.Naming","CA1710:IdentifiersShouldHaveCorrectSuffix")]
    public class MenuVM: AbstractContainer2VM<IMenuSource,IMenuVM>, IMenuVM,
            IActivatable<IMenuSource,IMenuVM>, IImageableVM {
        internal MenuVM(ViewModelFactory factory, string itemId, IEnumerable<IControlVM> controls)
        : base(itemId,controls) { }

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
    }
}
