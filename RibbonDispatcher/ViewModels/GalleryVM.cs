////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using Microsoft.Office.Core;

namespace PGSolutions.RibbonDispatcher.ViewModels {
    /// <summary>The ViewModel for Ribbon DropDown objects.</summary>
    internal class GalleryVM : AbstractControlVM<IGallerySource>, IGalleryVM,
            IActivatable<IGallerySource,IGalleryVM>, ISelectItemsVM {
        public GalleryVM(string itemId)
        : base(itemId) { }

        /// <inheritdoc/>
        public virtual string Description => (Strings as IControlStrings2)?.Description ?? $"{Id} Description";

        #region IActivatable implementation
        public new IGalleryVM Attach(IGallerySource source) => Attach<GalleryVM>(source);

        public override void Detach() { SelectionMade = null; base.Detach(); }
        #endregion

        #region IListable implementation
        /// <summary>Call back for ItemCount events from the drop-down ribbon elements.</summary>
        public int    ItemCount                => Source?.Count ?? 0;
        /// <summary>Call back for GetItemID events from the drop-down ribbon elements.</summary>
        public string ItemId(int Index)        => Source[Index].Id;
        /// <summary>Call back for GetItemLabel events from the drop-down ribbon elements.</summary>
        public string ItemLabel(int Index)     => Source[Index].Strings.Label;
        /// <summary>Call back for GetItemScreenTip events from the drop-down ribbon elements.</summary>
        public string ItemScreenTip(int Index) => Source[Index].Strings.ScreenTip;
        /// <summary>Call back for GetItemSuperTip events from the drop-down ribbon elements.</summary>
        public string ItemSuperTip(int Index)  => Source[Index].Strings.SuperTip;
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
