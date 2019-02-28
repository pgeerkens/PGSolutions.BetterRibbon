﻿////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using Microsoft.Office.Core;

namespace PGSolutions.RibbonDispatcher.ViewModels {
    internal class ComboBoxVM: AbstractControlVM<IComboBoxSource>, IComboBoxVM,
            IActivatable<IComboBoxSource, IComboBoxVM>, IEditableVM {
        public ComboBoxVM(string itemId) : base(itemId) { }

        #region IActivatable implementation
        public new IComboBoxVM Attach(IComboBoxSource source) => Attach<ComboBoxVM>(source);

        public override void Detach() { Edited = null; base.Detach(); }
        #endregion

        #region IListable implementation
        /// <summary>Call back for ItemCount events from the drop-down ribbon elements.</summary>
        public int ItemCount => Source?.Count ?? 0;

        /// <summary>.</summary>
        /// <param name="index">Index in the selection-list of the item being queried.</param>
        public IStaticItemVM this[int index] => Source[index];
        #endregion

        #region IEditable implementation
        public event EditedEventHandler Edited;

        public string Text => Source?.Text ?? "";

        public void OnEdited(IRibbonControl control, string text)
        => Edited?.Invoke(control, text);
        #endregion
    }
}
