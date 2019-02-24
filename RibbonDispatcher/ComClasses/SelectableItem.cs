////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////

using PGSolutions.RibbonDispatcher.ComInterfaces;
using Microsoft.Office.Core;

namespace PGSolutions.RibbonDispatcher.ComClasses.ViewModels {
    /// <summary>TODO</summary>
    internal class SelectableItemVM : AbstractControlVM<ISelectableItemSource>, ISelectableItemVM,
            IActivatable<ISelectableItemSource,ISelectableItemVM>, IClickableVM, IImageableVM {
        /// <summary>TODO</summary>
        internal SelectableItemVM(string ItemId) : base(ItemId) { }

        #region IActivatable implementation
        /// <inheritdoc/>
        public new ISelectableItemVM Attach(ISelectableItemSource source) => Attach<SelectableItemVM>(source);
        #endregion

        #region IClickable implementation
        /// <inheritdoc/>
        public event ClickedEventHandler Clicked;

        /// <inheritdoc/>
        public virtual void OnClicked(IRibbonControl control) => Clicked?.Invoke(control);
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
