////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

using Microsoft.Office.Core;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ComClasses.ViewModels;

namespace PGSolutions.RibbonDispatcher.ComClasses {
    /// <summary></summary>
    [SuppressMessage("Microsoft.Interoperability", "CA1409:ComVisibleTypesShouldBeCreatable")]
    [Description("")]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IEditedEvent))]
    [ComDefaultInterface(typeof(IEditBoxModel))]
    [Guid(Guids.EditBoxModel)]
    public class EditBoxModel : ControlModel<IEditBoxSource, IEditBoxVM>,
            IEditBoxModel, IEditBoxSource {
        internal EditBoxModel(Func<string, EditBoxVM> funcViewModel,
                IControlStrings strings, bool isEnabled, bool isVisible)
        : base(funcViewModel, strings, isEnabled, isVisible)
        { }

        public event EditedEventHandler Edited;

        public string Text { get; set; } = "";

        public IEditBoxModel Attach(string controlId) {
            ViewModel = AttachToViewModel(controlId, this);
            if (ViewModel != null) {
                ViewModel.Edited+= OnEdited;
                ViewModel.Invalidate();
            }
            return this;
        }

        private void OnEdited(IRibbonControl control, string text) => Edited?.Invoke(control,text);
    }
}
