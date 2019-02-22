////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Linq;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses.ViewModels {
    /// <summary>The ViewModel for Ribbon DropDown objects.</summary>
    public class DropDownVM : AbstractControlVM<IRibbonDropDownSource>, IRibbonDropDown,
            IActivatable<IRibbonDropDownSource,DropDownVM>, ISelectable {
        internal DropDownVM(string itemId)
        : base(itemId) { }

        #region IActivatable implementation
        public new DropDownVM Attach(IRibbonDropDownSource source) => Attach<DropDownVM>(source);

        public override void Detach() {
            SelectionMade = null;
            base.Detach();
        }
        #endregion

        #region ISelectable implementation
        /// <summary>TODO</summary>
        public event SelectedEventHandler  SelectionMade;

        /// <inheritdoc/>
        public string   SelectedItemId => Source[SelectedItemIndex].Id;

        /// <inheritdoc/>
        public int      SelectedItemIndex => Source?.SelectedIndex ?? 0;

        /// <summary>Call back for OnAction events from the drop-down ribbon elements.</summary>
        public void OnActionDropDown(string SelectedId, int SelectedIndex) {
            SelectionMade?.Invoke(this, SelectedIndex);
            Invalidate();
        }

        /// <inheritdoc/>
        public ISelectableItem this[int ItemIndex] => Source[ItemIndex];
        /// <inheritdoc/>
        public ISelectableItem this[string ItemId]
        => Source.FirstOrDefault(i => i.Id == ItemId);

        /// <summary>Call back for ItemCount events from the drop-down ribbon elements.</summary>
        public int      ItemCount                => Source?.Count ?? 0;
        /// <summary>Call back for GetItemID events from the drop-down ribbon elements.</summary>
        public string   ItemId(int Index)        => Source[Index].Id;
        /// <summary>Call back for GetItemLabel events from the drop-down ribbon elements.</summary>
        public string   ItemLabel(int Index)     => Source[Index].Label;
        /// <summary>Call back for GetItemScreenTip events from the drop-down ribbon elements.</summary>
        public string   ItemScreenTip(int Index) => Source[Index].ScreenTip;
        /// <summary>Call back for GetItemSuperTip events from the drop-down ribbon elements.</summary>
        public string   ItemSuperTip(int Index)  => Source[Index].SuperTip;
        /// <summary>Call back for GetItemLabel events from the drop-down ribbon elements.</summary>
        public object   ItemImage(int Index)     => "MacroSecurity";
        /// <summary>Call back for GetItemScreenTip events from the drop-down ribbon elements.</summary>
        public bool     ItemShowImage(int Index) => Source[Index].ShowImage;
        /// <summary>Call back for GetItemSuperTip events from the drop-down ribbon elements.</summary>
        public bool     ItemShowLabel(int Index) => Source[Index].ShowImage;
        #endregion
    }
}
