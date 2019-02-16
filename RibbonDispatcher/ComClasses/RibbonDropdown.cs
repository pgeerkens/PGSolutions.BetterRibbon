////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
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
