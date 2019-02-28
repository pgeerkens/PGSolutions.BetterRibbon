﻿////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using Microsoft.Office.Core;

namespace PGSolutions.RibbonDispatcher.ViewModels {
    /// <summary>The ViewModel for Ribbon DropDown objects.</summary>
    internal class DropDownVM : AbstractControlVM<IDropDownSource>, IDropDownVM,
            IActivatable<IDropDownSource,IDropDownVM>, ISelectItemsVM {
        public DropDownVM(string itemId)
        : base(itemId) { }

        #region IActivatable implementation
        public new IDropDownVM Attach(IDropDownSource source) => Attach<DropDownVM>(source);

        public override void Detach() { SelectionMade = null; base.Detach(); }
        #endregion

        #region IListable implementation
        /// <summary>Call back for ItemCount events from the drop-down ribbon elements.</summary>
        public int ItemCount => Source?.Count ?? 0;

        /// <summary>.</summary>
        /// <param name="index">Index in the selection-list of the item being queried.</param>
        public IStaticItemVM this[int index] => Source[index];
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
    }
}
