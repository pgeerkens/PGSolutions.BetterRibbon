////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using Microsoft.Office.Core;

using PGSolutions.RibbonDispatcher.ComInterfaces;

namespace PGSolutions.RibbonDispatcher.ComClasses.ViewModels {
    public class EditBoxVM : AbstractControlVM<IEditBoxSource>, IEditBox,
            IActivatable<IEditBoxSource,EditBoxVM>, IEditable {
        internal EditBoxVM(string itemId) : base(itemId) { }

        #region IActivatable implementation
        public new EditBoxVM Attach(IEditBoxSource source) => Attach<EditBoxVM>(source);

        public override void Detach() {
            Edited = null;
            base.Detach();
        }
        #endregion

        #region IEditable implementation
        public event EditedEventHandler Edited;

        public string Text => Source?.Text ?? "";

        public void OnEdited(IRibbonControl control, string text)
        => Edited?.Invoke(control, text);
        #endregion

    }
}
