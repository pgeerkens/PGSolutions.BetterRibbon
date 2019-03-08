////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System.Collections.Generic;
using Microsoft.Office.Core;

namespace PGSolutions.RibbonDispatcher.ViewModels {
    /// <summary>The ViewModel for static ribbon ComboBox objects.</summary>
    internal class StaticComboBoxVM: AbstractControlVM<IStaticComboBoxSource,IStaticComboBoxVM>, IStaticComboBoxVM,
            IActivatable<IStaticComboBoxSource,IStaticComboBoxVM>, IEditableVM {
        public StaticComboBoxVM(string itemId, IReadOnlyList<StaticItemVM> items)
        : base(itemId) => Items = items;

        #region IActivatable implementation
        public override IStaticComboBoxVM Attach(IStaticComboBoxSource source) => Attach<StaticComboBoxVM>(source);

        public override void Detach() { Edited = null; base.Detach(); }
        #endregion

        #region IListable implementation
        public IReadOnlyList<IStaticItemVM> Items { get; }
        #endregion

        #region IEditable implementation
        public event EditedEventHandler Edited;

        public string Text => Source?.Text ?? "";

        public void OnEdited(IRibbonControl control, string text)
        => Edited?.Invoke(control, text);
        #endregion
    }
}
