////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using Microsoft.Office.Core;

namespace PGSolutions.RibbonDispatcher.ViewModels {
    internal class EditBoxVM : AbstractControlVM<IEditBoxSource,IEditBoxVM>, IEditBoxVM,
            IActivatable<IEditBoxSource,IEditBoxVM>, IEditableVM {
        public EditBoxVM(string itemId) : base(itemId) { }

        #region IActivatable implementation
        public override IEditBoxVM Attach(IEditBoxSource source) => Attach<EditBoxVM>(source);

        public override void Detach() { Edited = null; base.Detach(); }
        #endregion

        #region IEditable implementation
        public event EditedEventHandler Edited;

        public string Text => Source?.Text ?? "";

        public void OnEdited(IRibbonControl control, string text)
        => Edited?.Invoke(control, text);
        #endregion

    }
}
