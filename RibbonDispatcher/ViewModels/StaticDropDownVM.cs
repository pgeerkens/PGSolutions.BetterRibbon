////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Collections.Generic;
using Microsoft.Office.Core;

namespace PGSolutions.RibbonDispatcher.ViewModels {
    /// <summary>The ViewModel for static ribbon DropDown objects.</summary>
    internal class StaticDropDownVM : AbstractControlVM<IStaticDropDownSource,IDropDownVM>, IDropDownVM,
            IActivatable<IStaticDropDownSource,IDropDownVM>, ISelectItemsVM {
        internal StaticDropDownVM(string itemId, IReadOnlyList<StaticItemVM> items)
        : base(itemId) => Items = items;

        #region IActivatable implementation
        public override IDropDownVM Attach(IStaticDropDownSource source) => Attach<StaticDropDownVM>(source);

        public override void Detach() { SelectionMade = null; base.Detach(); }
        #endregion

        #region IListable implementation
        public IReadOnlyList<IStaticItemVM> Items { get; }
        #endregion

        #region ISelectable implementation
        /// <summary>TODO</summary>
        public event SelectionMadeEventHandler  SelectionMade;

        /// <inheritdoc/>
        public string   SelectedItemId    => Source?.SelectedId ?? "";

        /// <inheritdoc/>
        public int      SelectedItemIndex => Source?.SelectedIndex ?? 0;

        /// <summary>Call back for OnAction events from the drop-down ribbon elements.</summary>
        public void OnSelectionMade(IRibbonControl control, string selectedId, int selectedIndex) {
            SelectionMade?.Invoke(control, selectedId, selectedIndex);
            Invalidate();
        }
        #endregion
    }
}
