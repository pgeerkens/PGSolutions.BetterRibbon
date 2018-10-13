////////////////////////////////////////////////////////////////////////////////////////////////////
//                                Copyright (c) 2018 Pieter Geerkens                              //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.ControlMixins;
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
        ISelectableMixin {
        internal RibbonDropDown(string itemId, IRibbonControlStrings strings, bool visible, bool enabled
        ) : base(itemId, strings, visible, enabled) { }

        #region IActivatable implementation
        private bool _isAttached    = false;

        public override bool IsEnabled => base.IsEnabled && _isAttached;
        public override bool IsVisible => base.IsVisible || ShowWhenInactive;

        public bool ShowWhenInactive { get; set; } = true;

        private Func<int> Getter { get; set; }

        public IRibbonDropDown Attach(Func<int> getter) {
            _isAttached = true;
            Getter = getter;
            return this;
        }

        public void Detach() {
            Getter = ()=>0;
            _isAttached = false;
            SetLanguageStrings(RibbonTextLanguageControl.Empty);
        }

        IRibbonCommon IActivatableControl<IRibbonCommon, int>.Attach(Func<int> getter) =>
            Attach(getter) as IRibbonCommon;
        void IActivatableControl<IRibbonCommon, int>.Detach() => Detach();
        #endregion

        /// <summary>TODO</summary>
        public event SelectedEventHandler SelectionMade;

        private IList<ISelectableItem>  _items  = new List<ISelectableItem>();

        /// <inheritdoc/>
        public string   SelectedItemId => _items[SelectedItemIndex].Id;

        /// <inheritdoc/>
        public int      SelectedItemIndex => Getter();

        /// <summary>Call back for OnAction events from the drop-down ribbon elements.</summary>
        public void OnActionDropDown(string SelectedId, int SelectedIndex) {
            //_selectedItemIndex = SelectedIndex;
            SelectionMade?.Invoke(SelectedId, SelectedIndex);
            OnChanged();
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
    }
}
