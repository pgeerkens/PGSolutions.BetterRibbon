////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ViewModels;

namespace PGSolutions.RibbonDispatcher.Models {
    /// <summary>The COM visible Model for Ribbon ComboBox controls.</summary>
    [Description("The COM visible Model for Ribbon ComboBox controls.")]
    [CLSCompliant(true)]
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IEditedEvent))]
    [ComDefaultInterface(typeof(IComboBoxModel))]
    [Guid(Guids.ComboBoxModel)]
    public sealed class ComboBoxModel: ControlModel<IComboBoxSource,IComboBoxVM>, IComboBoxModel,
            IComboBoxSource {
        internal ComboBoxModel(Func<string, ComboBoxVM> funcViewModel, IControlStrings strings)
        : base(funcViewModel, strings) { }

        #region IActivatable implementation
        public IComboBoxModel Attach(string controlId) {
            ViewModel = AttachToViewModel(controlId, this);
            if (ViewModel != null) { ViewModel.Edited += OnEdited; }
            return this;
        }

        public override void Detach() { Edited = null;  base.Detach(); }
        #endregion

        #region IImageable implementation
        public IImageObject Image     { get; set; } = "MacroSecurity".ToImageObject();
        public bool         ShowImage { get; set; } = true;
        public bool         ShowLabel { get; set; } = true;

        public IComboBoxModel SetImage(IImageObject image) { Image = image; return this; }
        #endregion

        #region IListable implementation
        public IReadOnlyList<IStaticItemVM> Items => _items.AsReadOnly();
        #endregion

        #region IDynamicListable implementation
        private List<IStaticItemVM> _items = new List<IStaticItemVM>();

        public IComboBoxModel ClearList() { _items.Clear(); return this; }

        public IComboBoxModel AddSelectableModel(IStaticItemVM selectableModel) {
            _items.Add(selectableModel);
            ViewModel?.Invalidate();
            return this;
        }
        #endregion

        #region IEditableList implementation
        public event EditedEventHandler Edited;

        public string Text { get; set; } = "";

        private void OnEdited(IRibbonControl control, string text) => Edited?.Invoke(control, text);
        #endregion
    }
}
