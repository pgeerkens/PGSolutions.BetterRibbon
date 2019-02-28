////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Collections.Generic;
using Microsoft.Office.Core;

namespace PGSolutions.RibbonDispatcher.ViewModels {
    /// <summary>The ViewModel for Ribbon DropDown objects.</summary>
    internal class StaticGalleryVM : AbstractControlVM<IStaticGallerySource>, IGalleryVM,
            IActivatable<IStaticGallerySource,IGalleryVM>, ISelectItemsVM {
        public StaticGalleryVM(string itemId, IList<StaticItemVM> items)
        : base(itemId) => Items = items;

        private IList<StaticItemVM> Items { get; }

        /// <inheritdoc/>
        public virtual string Description => (Strings as IControlStrings2)?.Description ?? $"{Id} Description";

        #region IActivatable implementation
        public new IGalleryVM Attach(IStaticGallerySource source) => Attach<StaticGalleryVM>(source);

        public override void Detach() { SelectionMade = null; base.Detach(); }
        #endregion

        #region IListable implementation
        /// <summary>Call back for ItemCount events from the drop-down ribbon elements.</summary>
        public int    ItemCount                => Items?.Count ?? 0;
        /// <summary>Call back for GetItemID events from the drop-down ribbon elements.</summary>
        public string ItemId(int Index)        => Items[Index].Id;
        /// <summary>Call back for GetItemLabel events from the drop-down ribbon elements.</summary>
        public string ItemLabel(int Index)     => Items[Index].Label;
        /// <summary>Call back for GetItemScreenTip events from the drop-down ribbon elements.</summary>
        public string ItemScreenTip(int Index) => Items[Index].ScreenTip;
        /// <summary>Call back for GetItemSuperTip events from the drop-down ribbon elements.</summary>
        public string ItemSuperTip(int Index)  => Items[Index].SuperTip;
        /// <summary>Call back for GetItemLabel events from the drop-down ribbon elements.</summary>
        public object ItemImage(int Index)     => "MacroSecurity";
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
