////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Collections.Generic;
using Microsoft.Office.Core;

namespace PGSolutions.RibbonDispatcher.ViewModels {
    /// <summary>The ViewModel for static ribbon ComboBox objects.</summary>
    internal class StaticComboBoxVM: AbstractControlVM<IStaticComboBoxSource>, IComboBoxVM,
            IActivatable<IStaticComboBoxSource,IComboBoxVM>, IEditableVM {
        public StaticComboBoxVM(string itemId, IList<StaticItemVM> items)
        : base(itemId) => Items = items;

        private IList<StaticItemVM> Items { get; }

        #region IActivatable implementation
        public new IComboBoxVM Attach(IStaticComboBoxSource source) => Attach<StaticComboBoxVM>(source);

        public override void Detach() { Edited = null; base.Detach(); }
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
        /// <summary>Call back for GetItemScreenTip events from the drop-down ribbon elements.</summary>
        public bool   ItemShowImage(int Index) => Items[Index].ShowImage;
        /// <summary>Call back for GetItemSuperTip events from the drop-down ribbon elements.</summary>
        public bool   ItemShowLabel(int Index) => Items[Index].ShowImage;
        #endregion

        #region IEditable implementation
        public event EditedEventHandler Edited;

        public string Text => Source?.Text ?? "";

        public void OnEdited(IRibbonControl control, string text)
        => Edited?.Invoke(control, text);
        #endregion
    }
}
