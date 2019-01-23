////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.Utilities;

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
    public class RibbonDropDown : RibbonCommon, IRibbonDropDown, IActivatableControl<IRibbonCommon, int>,
        ISelectable {
        internal RibbonDropDown(string itemId, IRibbonControlStrings strings, bool visible, bool enabled,
            Func<int> getter = null
        ) : base(itemId, strings, visible, enabled) {
            if (getter != null) { Attach(getter); }
        }

        #region IActivatable implementation
        private Func<int> Getter { get; set; }

        public IRibbonDropDown Attach(Func<int> getter) {
            base.Attach();
            Getter = getter;
            return this;
        }

        public override void Detach() {
            Getter = ()=>0;
            SelectionMade = null;
            base.Detach();
        }

        IRibbonCommon IActivatableControl<IRibbonCommon, int>.Attach(Func<int> getter) =>
            Attach(getter) as IRibbonCommon;
        void IActivatableControl<IRibbonCommon, int>.Detach() => Detach();
        #endregion

        #region ISelectable implementation
        /// <summary>TODO</summary>
        public event SelectedEventHandler SelectionMade;

        private IList<ISelectableItem>  _items  = new List<ISelectableItem>();

        /// <inheritdoc/>
        public string   SelectedItemId => _items[SelectedItemIndex].Id;

        /// <inheritdoc/>
        public int      SelectedItemIndex => Getter?.Invoke() ?? 0;

        /// <summary>Call back for OnAction events from the drop-down ribbon elements.</summary>
        public void OnActionDropDown(string SelectedId, int SelectedIndex) {
            SelectionMade?.Invoke(SelectedId, SelectedIndex);
            Invalidate();
        }

        /// <summary>Returns this RibbonDropDown with a new {SelectableItem} in its list.</summary>
        public IRibbonDropDown AddItem(ISelectableItem SelectableItem) {
            _items.Add(SelectableItem);
            return this;
        }

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
