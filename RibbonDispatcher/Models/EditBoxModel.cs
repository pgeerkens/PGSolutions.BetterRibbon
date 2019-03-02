////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

using Microsoft.Office.Core;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.RibbonDispatcher.Models {
    /// <summary>The COM visible Model for Ribbon EditBox controls.</summary>
    [Description("The COM visible Model for Ribbon EditBox controls.")]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IEditedEvent))]
    [ComDefaultInterface(typeof(IEditBoxModel))]
    [Guid(Guids.EditBoxModel)]
    public class EditBoxModel : ControlModel<IEditBoxSource, IEditBoxVM>,
            IEditBoxModel, IEditBoxSource {
        internal EditBoxModel(Func<string, EditBoxVM> funcViewModel,
                IControlStrings strings)
        : base(funcViewModel, strings)
        { }

        public IEditBoxModel Attach(string controlId) {
            ViewModel = AttachToViewModel(controlId, this);
            if (ViewModel != null) { ViewModel.Edited+= OnEdited; }
            return this;
        }

        #region IEditable implementation
        public event EditedEventHandler Edited;

        public string Text { get; set; } = "";

        private void OnEdited(IRibbonControl control, string text) => Edited?.Invoke(control,text);
        #endregion

        #region IImageable implementation
        public IImageObject Image     { get; set; } = "MacroSecurity".ToImageObject();
        public bool         ShowImage { get; set; } = true;
        public bool         ShowLabel { get; set; } = true;

        public IEditBoxModel SetImage(IImageObject image) { Image = image; return this; }
        #endregion
    }
}
