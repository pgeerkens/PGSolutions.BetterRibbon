////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ComClasses.ViewModels;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    /// <summary></summary>
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable")]
    [Description("")]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IEditedEvents))]
    [ComDefaultInterface(typeof(IEditBoxModel))]
    [Guid(Guids.EditBoxModel)]
    public class EditBoxModel : RibbonControlModel<IEditBoxSource, EditBoxVM>,
            IEditBoxModel, IEditBoxSource {
        public EditBoxModel(Func<string, EditBoxVM> funcViewModel,
                IControlStrings strings, bool isEnabled, bool isVisible)
        : base(funcViewModel, strings, isEnabled, isVisible)
        { }

        public event EditedEventHandler Edited;

        public string Text { get; } = "";

        public IEditBoxModel Attach(string controlId) {
            ViewModel = AttachToViewModel(controlId, this);
            if (ViewModel != null) {
                ViewModel.Edited+= OnEdited;
                ViewModel.Invalidate();
            }
            return this;
        }

        private void OnEdited(object sender, string text) => Edited?.Invoke(sender,text);
    }
}
