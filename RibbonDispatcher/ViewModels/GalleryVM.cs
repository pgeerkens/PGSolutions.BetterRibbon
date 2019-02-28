////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Collections.Generic;
using Microsoft.Office.Core;

namespace PGSolutions.RibbonDispatcher.ViewModels {
    /// <summary>The ViewModel for Ribbon DropDown objects.</summary>
    internal class GalleryVM : AbstractControlVM<IGallerySource>, IGalleryVM,
            IActivatable<IGallerySource,IGalleryVM> {
        public GalleryVM(string itemId)
        : base(itemId) { }

        /// <inheritdoc/>
        public virtual string Description => (Strings as IControlStrings2)?.Description ?? $"{Id} Description";

        #region IActivatable implementation
        public new IGalleryVM Attach(IGallerySource source) => Attach<GalleryVM>(source);

        public override void Detach() { SelectionMade = null; base.Detach(); }
        #endregion

        #region IListable implementation
        public IReadOnlyList<IStaticItemVM> Items => Source?.Items;
        #endregion

        #region ISelectable implementation
        /// <summary>TODO</summary>
        public event SelectionMadeEventHandler  SelectionMade;

        /// <inheritdoc/>
        public string SelectedItemId    => Source?.SelectedId ?? "";

        /// <inheritdoc/>
        public int    SelectedItemIndex => Source?.SelectedIndex ?? 0;

        /// <summary>Call back for OnAction events from the drop-down ribbon elements.</summary>
        public void OnSelectionMade(IRibbonControl control, string selectedId, int selectedIndex) {
            SelectionMade?.Invoke(control, selectedId, selectedIndex);
            Invalidate();
        }
        #endregion

        #region IGallerySizeVM implementation
        public int ItemHeight => Source?.ItemHeight ?? default;
        public int ItemWidth  => Source?.ItemWidth  ?? default;
        #endregion
    }
}
