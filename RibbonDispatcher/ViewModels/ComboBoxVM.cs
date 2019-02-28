////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Collections.Generic;
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
        public IReadOnlyList<IStaticItemVM> Items => Source?.Items;
        #endregion

        #region IEditable implementation
        public event EditedEventHandler Edited;

        public string Text => Source?.Text ?? "";

        public void OnEdited(IRibbonControl control, string text)
        => Edited?.Invoke(control, text);
        #endregion
    }
}
