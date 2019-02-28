////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Collections.Generic;
using Microsoft.Office.Core;

namespace PGSolutions.RibbonDispatcher.ViewModels {
    /// <summary>The ViewModel for Ribbon DropDown objects.</summary>
    internal class StaticGalleryVM : AbstractControlVM<IStaticGallerySource>, IStaticGalleryVM,
            IActivatable<IStaticGallerySource,IStaticGalleryVM>, ISelectItemsVM {
        public StaticGalleryVM(string itemId, IReadOnlyList<StaticItemVM> items)
        : base(itemId) => Items = items;

        public IReadOnlyList<StaticItemVM> Items { get; }

        /// <inheritdoc/>
        public virtual string Description => (Strings as IControlStrings2)?.Description ?? $"{Id} Description";

        #region IActivatable implementation
        public new IStaticGalleryVM Attach(IStaticGallerySource source) => Attach<StaticGalleryVM>(source);

        public override void Detach() { SelectionMade = null; base.Detach(); }
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

        #region IListable implementation
        /// <summary>Call back for ItemCount events from the drop-down ribbon elements.</summary>
        public int    ItemCount                => Items?.Count ?? 0;

        /// <summary>.</summary>
        /// <param name="index">Index in the selection-list of the item being queried.</param>
        public IStaticItemVM this[int index] => Items[index];
        #endregion

        #region IGallerySizeVM implementation
        public int ItemHeight => Source?.ItemHeight ?? default;
        public int ItemWidth  => Source?.ItemWidth  ?? default;
        #endregion
    }
}
