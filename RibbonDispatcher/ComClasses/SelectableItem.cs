////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ComClasses.ViewModels;
using Microsoft.Office.Core;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    /// <summary>TODO</summary>
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable",
       Justification = "Public, Non-Creatable, class with exported Events.")]
    [Serializable]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ISelectableItem))]
    [Guid(Guids.SelectableItem)]
    internal class SelectableItem : AbstractControlVM<ISelectableItemSource>, ISelectableItem,
            IActivatable<ISelectableItemSource,SelectableItem>, IClickableVM, IImageableVM {
        /// <summary>TODO</summary>
        internal SelectableItem(string ItemId) : base(ItemId) { }

        #region IActivatable implementation
        /// <inheritdoc/>
        public new SelectableItem Attach(ISelectableItemSource source) => Attach<SelectableItem>(source);

        /// <inheritdoc/>
        public override void Detach() => base.Detach();
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
