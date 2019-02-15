////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    /// <summary>The ViewModel for Ribbon DropDown objects.</summary>
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable",
      Justification = "Public, Non-Creatable, class with exported Events.")]
    [Serializable]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(ISelectionMadeEvents))]
    [ComDefaultInterface(typeof(IRibbonDropDown))]
    [Guid(Guids.RibbonDropDown)]
    public class RibbonDropDown : RibbonCommon<IRibbonDropDownSource>, IRibbonDropDown,
            IActivatable<RibbonDropDown, IRibbonDropDownSource>, ISelectable {
        internal RibbonDropDown(string itemId)
        : base(itemId) { }

        #region IActivatable implementation
        RibbonDropDown IActivatable<RibbonDropDown, IRibbonDropDownSource>.Attach(IRibbonDropDownSource source)
        => Attach<RibbonDropDown>(source);

        public override void Detach() {
            SelectionMade = null;
            base.Detach();
        }
        #endregion

        #region ISelectable implementation
        /// <summary>TODO</summary>
        public event SelectedEventHandler  SelectionMade;

        private IList<ISelectableItem>  _items  => Source?.Items ?? new List<ISelectableItem>();

        /// <inheritdoc/>
        public string   SelectedItemId => _items[SelectedItemIndex].Id;

        /// <inheritdoc/>
        public int      SelectedItemIndex => Source?.SelectedIndex ?? 0;

        /// <summary>Call back for OnAction events from the drop-down ribbon elements.</summary>
        public void OnActionDropDown(string SelectedId, int SelectedIndex) {
            SelectionMade?.Invoke(this, SelectedIndex);
            Invalidate();
        }

        ///// <summary>Returns this RibbonDropDown with a new {SelectableItem} in its list.</summary>
        //public IRibbonDropDown AddItem(ISelectableItem SelectableItem) {
        //    _items.Add(SelectableItem);
        //    Invalidate();
        //    return this;
        //}

        /// <inheritdoc/>
        public ISelectableItem this[int ItemIndex] => _items[ItemIndex];
        /// <inheritdoc/>
        public ISelectableItem this[string ItemId] => ( from i in _items where i.Id == ItemId select i ).FirstOrDefault();

        /// <summary>Call back for ItemCount events from the drop-down ribbon elements.</summary>
        public int      ItemCount                => _items?.Count ?? -1;
        /// <summary>Call back for GetItemID events from the drop-down ribbon elements.</summary>
        public string   ItemId(int Index)        => _items[Index].Id;
        /// <summary>Call back for GetItemLabel events from the drop-down ribbon elements.</summary>
        public string   ItemLabel(int Index)     => _items[Index].Label;
        /// <summary>Call back for GetItemScreenTip events from the drop-down ribbon elements.</summary>
        public string   ItemScreenTip(int Index) => _items[Index].ScreenTip;
        /// <summary>Call back for GetItemSuperTip events from the drop-down ribbon elements.</summary>
        public string   ItemSuperTip(int Index)  => _items[Index].SuperTip;
        /// <summary>Call back for GetItemLabel events from the drop-down ribbon elements.</summary>
        public object   ItemImage(int Index)     => "MacroSecurity";
        /// <summary>Call back for GetItemScreenTip events from the drop-down ribbon elements.</summary>
        public bool     ItemShowImage(int Index) => _items[Index].ShowImage;
        /// <summary>Call back for GetItemSuperTip events from the drop-down ribbon elements.</summary>
        public bool     ItemShowLabel(int Index) => _items[Index].ShowImage;
        #endregion
    }
}
