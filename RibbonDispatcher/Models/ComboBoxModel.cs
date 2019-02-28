////////////////////////////////////////////////////////////////////////////////////////////////////
//                             Copyright (c) 2017-2019 Pieter Geerkens                            //
////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Runtime.InteropServices;

using PGSolutions.RibbonDispatcher.ComInterfaces;
using PGSolutions.RibbonDispatcher.ViewModels;
using Microsoft.Office.Core;

namespace PGSolutions.RibbonDispatcher.Models {
    /// <summary>The COM visible Model for Ribbon ComboBox controls.</summary>
    [SuppressMessage("Microsoft.Naming","CA1710:IdentifiersShouldHaveCorrectSuffix")]
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

        public IComboBoxModel Attach(string controlId) {
            ViewModel = AttachToViewModel(controlId, this);
            if (ViewModel != null) {
                ViewModel.Edited += OnEdited;
                ViewModel.Invalidate();
            }
            return this;
        }

        #region IListable implementation
        public IReadOnlyList<IStaticItemVM> Items => _items.AsReadOnly();
        private List<IStaticItemVM> _items = new List<IStaticItemVM>();

        public int FindId(string id)
        => Items.Where((i,n) => i.Id == id).Select((i,n)=>n).FirstOrDefault();
        #endregion

        #region IDynamicListable implementation
        public IComboBoxModel ClearList() { _items.Clear(); return this; }

        public IComboBoxModel AddSelectableModel(IStaticItemVM selectableModel) {
            _items.Add(selectableModel);
            ViewModel?.Invalidate();
            return this;
        }
        #endregion

        #region IEditableList implementation
        public event EditedEventHandler Edited;

        public string Text          { get; set; } = "";

        private void OnEdited(IRibbonControl control, string text) => Edited?.Invoke(control, text);
        #endregion
    }
}
